
import { NextRequest, NextResponse } from "next/server";
import { exec } from "child_process";
import fs from "fs";
import path from "path";
import { promisify } from "util";
import os from "os";

const execPromise = promisify(exec);
const writeFilePromise = promisify(fs.writeFile);
const readFilePromise = promisify(fs.readFile);
const unlinkPromise = promisify(fs.unlink);

export async function POST(req: NextRequest) {
    try {
        const { markdown } = await req.json();

        if (!markdown) {
            return NextResponse.json({ error: "Markdown content is required" }, { status: 400 });
        }

        // Create a unique temporary directory or file prefix
        const uniqueId = Math.random().toString(36).substring(7);
        const tempDir = os.tmpdir();
        const inputPath = path.join(tempDir, `input-${uniqueId}.md`);
        const outputPath = path.join(tempDir, `output-${uniqueId}.hml`);

        // Path to the python script
        // We assume the script is in web/scripts/md2hml.py relative to the project root
        // In production, this path might need adjustment involving process.cwd()
        const scriptPath = path.join(process.cwd(), "scripts", "md2hml.py");

        // Write input markdown to temp file
        await writeFilePromise(inputPath, markdown);

        // Execute python script
        try {
            await execPromise(`python3 "${scriptPath}" "${inputPath}" "${outputPath}"`);
        } catch (error: any) {
            console.error("Conversion execution error:", error);
            return NextResponse.json({ error: "Conversion failed execution", details: error.message }, { status: 500 });
        }

        // Read the output HML
        let hmlContent;
        try {
            hmlContent = await readFilePromise(outputPath, "utf-8");
        } catch (readError) {
            console.error("Error reading output:", readError);
            return NextResponse.json({ error: "Failed to read generated HML file" }, { status: 500 });
        }

        // Cleanup temp files
        try {
            await unlinkPromise(inputPath);
            await unlinkPromise(outputPath);
        } catch (cleanupError) {
            console.warn("Cleanup error:", cleanupError);
        }

        // Return the HML content
        // We send it as text/xml
        return new NextResponse(hmlContent, {
            status: 200,
            headers: {
                "Content-Type": "application/xml",
                "Content-Disposition": `attachment; filename="converted.hml"`,
            },
        });

    } catch (error: any) {
        console.error("API Error:", error);
        return NextResponse.json({ error: "Internal Server Error", details: error.message }, { status: 500 });
    }
}
