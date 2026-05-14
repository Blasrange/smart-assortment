import {
  AngularNodeAppEngine,
  createNodeRequestHandler,
  isMainModule,
  writeResponseToNodeResponse,
} from '@angular/ssr/node';
import { randomBytes } from 'node:crypto';
import { spawn } from 'node:child_process';
import { unlink, writeFile } from 'node:fs/promises';
import express from 'express';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const browserDistFolder = join(import.meta.dirname, '../browser');

const app = express();
const angularApp = new AngularNodeAppEngine();

app.use(express.json({ limit: '20mb' }));

function runPowerShell(script: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const ps = spawn('powershell', ['-NoProfile', '-NonInteractive', '-Command', script]);
    let out = '';
    let err = '';

    ps.stdout?.on('data', (data: Buffer) => (out += data.toString()));
    ps.stderr?.on('data', (data: Buffer) => (err += data.toString()));
    ps.on('error', reject);
    ps.on('close', (code: number | null) => {
      if (code !== 0) {
        reject(new Error(err.trim() || `Exit ${code}`));
        return;
      }
      resolve(out.trim());
    });

    const timer = setTimeout(() => {
      ps.kill();
      reject(new Error('Timeout de PowerShell'));
    }, 60000);

    ps.on('close', () => clearTimeout(timer));
  });
}

function buildRawPrintScript(escapedPrinterName: string, zplFilePath: string): string {
  return `
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
public class DocInfo {
  [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
  [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
  [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
}
public class RawPrint {
  [DllImport("winspool.Drv", EntryPoint="OpenPrinterA", SetLastError=true, CharSet=CharSet.Ansi)]
  public static extern bool OpenPrinter(string n, out IntPtr h, IntPtr d);
  [DllImport("winspool.Drv", EntryPoint="ClosePrinter", SetLastError=true)]
  public static extern bool ClosePrinter(IntPtr h);
  [DllImport("winspool.Drv", EntryPoint="StartDocPrinterA", SetLastError=true, CharSet=CharSet.Ansi)]
  public static extern int StartDocPrinter(IntPtr h, int lv, [In, MarshalAs(UnmanagedType.LPStruct)] DocInfo d);
  [DllImport("winspool.Drv", EntryPoint="EndDocPrinter", SetLastError=true)]
  public static extern bool EndDocPrinter(IntPtr h);
  [DllImport("winspool.Drv", EntryPoint="StartPagePrinter", SetLastError=true)]
  public static extern bool StartPagePrinter(IntPtr h);
  [DllImport("winspool.Drv", EntryPoint="EndPagePrinter", SetLastError=true)]
  public static extern bool EndPagePrinter(IntPtr h);
  [DllImport("winspool.Drv", EntryPoint="WritePrinter", SetLastError=true)]
  public static extern bool WritePrinter(IntPtr h, IntPtr p, int c, out int w);
}
"@

$printerName = '${escapedPrinterName}'
$zplPath     = '${zplFilePath}'
$bytes = [System.IO.File]::ReadAllBytes($zplPath)

$hPrinter = [IntPtr]::Zero
if (-not [RawPrint]::OpenPrinter($printerName, [ref]$hPrinter, [IntPtr]::Zero)) {
  throw "No se pudo abrir la impresora: $printerName"
}

$di = New-Object DocInfo
$di.pDocName    = 'ZPL Label'
$di.pOutputFile = $null
$di.pDataType   = 'RAW'

if ([RawPrint]::StartDocPrinter($hPrinter, 1, $di) -le 0) {
  [RawPrint]::ClosePrinter($hPrinter) | Out-Null
  throw 'StartDocPrinter falló'
}

[RawPrint]::StartPagePrinter($hPrinter) | Out-Null
$ptr     = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($bytes.Length)
[System.Runtime.InteropServices.Marshal]::Copy($bytes, 0, $ptr, $bytes.Length)
$written = 0
$ok      = [RawPrint]::WritePrinter($hPrinter, $ptr, $bytes.Length, [ref]$written)
[System.Runtime.InteropServices.Marshal]::FreeHGlobal($ptr)
[RawPrint]::EndPagePrinter($hPrinter)  | Out-Null
[RawPrint]::EndDocPrinter($hPrinter)   | Out-Null
[RawPrint]::ClosePrinter($hPrinter)    | Out-Null

if (-not $ok) { throw 'WritePrinter falló — verifique que la impresora acepta datos RAW' }
Write-Output "OK:$written"
`;
}

app.get('/api/exito-labels/printers', async (_req, res) => {
  try {
    const output = await runPowerShell(
      "Get-WmiObject -Query 'SELECT Name,Status FROM Win32_Printer' | Select-Object Name,Status | ConvertTo-Json",
    );

    if (!output) {
      res.json({ printers: [] });
      return;
    }

    const raw = JSON.parse(output);
    const list = Array.isArray(raw) ? raw : [raw];

    res.json({
      printers: list
        .filter((printer) => printer?.Name)
        .map((printer) => ({
          name: String(printer.Name),
          status: String(printer.Status ?? ''),
        })),
    });
  } catch {
    res.status(500).json({ error: 'No se pudieron obtener las impresoras.' });
  }
});

app.post('/api/exito-labels/print', async (req, res) => {
  let tmpFile: string | null = null;

  try {
    const { printerName, zpl } = req.body as { printerName?: unknown; zpl?: unknown };

    if (typeof printerName !== 'string' || printerName.trim().length === 0 || printerName.length > 256) {
      res.status(400).json({ error: 'Nombre de impresora inválido.' });
      return;
    }

    if (typeof zpl !== 'string' || zpl.trim().length === 0) {
      res.status(400).json({ error: 'Contenido ZPL inválido o vacío.' });
      return;
    }

    if (zpl.length > 20_000_000) {
      res.status(400).json({ error: 'El ZPL supera el límite permitido (20 MB).' });
      return;
    }

    const tmpName = `zpl_${randomBytes(8).toString('hex')}.prn`;
    tmpFile = join(tmpdir(), tmpName);
    await writeFile(tmpFile, zpl, 'ascii');

    const escapedName = printerName.replace(/'/g, "''");
    const script = buildRawPrintScript(escapedName, tmpFile);
    const result = await runPowerShell(script);

    res.json({ success: true, message: result });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Error desconocido';
    res.status(500).json({ error: `Error al imprimir: ${message}` });
  } finally {
    if (tmpFile) {
      unlink(tmpFile).catch(() => undefined);
    }
  }
});

/**
 * Example Express Rest API endpoints can be defined here.
 * Uncomment and define endpoints as necessary.
 *
 * Example:
 * ```ts
 * app.get('/api/{*splat}', (req, res) => {
 *   // Handle API request
 * });
 * ```
 */

/**
 * Serve static files from /browser
 */
app.use(
  express.static(browserDistFolder, {
    maxAge: '1y',
    index: false,
    redirect: false,
  }),
);

/**
 * Handle all other requests by rendering the Angular application.
 */
app.use((req, res, next) => {
  angularApp
    .handle(req)
    .then((response) =>
      response ? writeResponseToNodeResponse(response, res) : next(),
    )
    .catch(next);
});

/**
 * Start the server if this module is the main entry point, or it is ran via PM2.
 * The server listens on the port defined by the `PORT` environment variable, or defaults to 4000.
 */
if (isMainModule(import.meta.url) || process.env['pm_id']) {
  const port = process.env['PORT'] || 4000;
  app.listen(port, (error) => {
    if (error) {
      throw error;
    }

    console.log(`Node Express server listening on http://localhost:${port}`);
  });
}

/**
 * Request handler used by the Angular CLI (for dev-server and during build) or Firebase Cloud Functions.
 */
export const reqHandler = createNodeRequestHandler(app);
