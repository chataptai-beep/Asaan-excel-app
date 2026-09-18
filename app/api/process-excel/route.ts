import { NextRequest, NextResponse } from "next/server";

// Allow up to 20 minutes for local dev processing of large workbooks
export const maxDuration = 1200;
import { writeFileSync, readFileSync, unlinkSync, existsSync } from "fs";
import { join } from "path";
import { tmpdir } from "os";
import { spawn } from "child_process";
import { randomBytes } from "crypto";

function runPython(script: string, input: string, output: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const child = spawn("python", [script, "--input", input, "--output", output]);

    let out = "";
    let err = "";
    child.stdout.on("data", (d: Buffer) => { out += d.toString(); });
    child.stderr.on("data", (d: Buffer) => { err += d.toString(); });

    const timer = setTimeout(() => {
      child.kill();
      reject(new Error("Timed out after 20 minutes"));
    }, 20 * 60 * 1000);

    child.on("close", (code: number | null) => {
      clearTimeout(timer);
      if (code === 0) resolve(out);
      else reject(new Error((err || out || `exit ${code}`).slice(0, 800)));
    });

    child.on("error", (e: Error) => { clearTimeout(timer); reject(e); });
  });
}

export async function POST(req: NextRequest) {
  let inputPath = "";
  let outputPath = "";

  try {
    const formData = await req.formData();
    const file = formData.get("file") as File | null;
    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }

    const id = randomBytes(8).toString("hex");
    inputPath  = join(tmpdir(), `xl_in_${id}.xlsx`);
    outputPath = join(tmpdir(), `xl_out_${id}.xlsx`);
    const scriptPath = join(process.cwd(), "scripts", "process_excel.py");

    const buffer = Buffer.from(await file.arrayBuffer());
    writeFileSync(inputPath, buffer);

    await runPython(scriptPath, inputPath, outputPath);

    if (!existsSync(outputPath)) {
      return NextResponse.json({ error: "Output file was not created" }, { status: 500 });
    }

    const outBuffer = readFileSync(outputPath);
    const baseName  = file.name.replace(/\.[^/.]+$/, "");

    return new NextResponse(outBuffer, {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${baseName}_clean.xlsx"`,
      },
    });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: msg }, { status: 500 });
  } finally {
    try { if (inputPath)  unlinkSync(inputPath);  } catch {}
    try { if (outputPath) unlinkSync(outputPath); } catch {}
  }
}
