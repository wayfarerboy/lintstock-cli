const fs = require("node:fs");
const path = require("node:path");
const pdf = require("pdf-parse");
const chalk = require("chalk");
const cliProgress = require("cli-progress");

// This script will take two arguments:
// --input-dir: the directory to read PDFs from
// --output-dir: the directory to write Markdown files to

async function convertPdfToMarkdown(pdfPath) {
  const dataBuffer = fs.readFileSync(pdfPath);
  const data = await pdf(dataBuffer);
  return data.text;
}

// --- Argument Parsing ---
const args = process.argv.slice(2);
let inputDir = null;
let outputDir = null;

for (let i = 0; i < args.length; i++) {
  if (args[i] === "--input-dir") {
    const dir = args[i + 1];
    if (dir && !dir.startsWith("-")) {
      inputDir = dir;
      i++; // Skip the value
    }
  } else if (args[i] === "--output-dir") {
    const dir = args[i + 1];
    if (dir && !dir.startsWith("-")) {
      outputDir = dir;
      i++; // Skip the value
    }
  }
}

if (!inputDir || !outputDir) {
  console.error(
    chalk.red("Both --input-dir and --output-dir must be provided."),
  );
  process.exit(1);
}

// --- File Processing ---
let filePaths = [];
try {
  if (fs.existsSync(inputDir) && fs.statSync(inputDir).isDirectory()) {
    const allFiles = fs.readdirSync(inputDir);
    filePaths = allFiles
      .filter((file) => file.toLowerCase().endsWith(".pdf"))
      .map((file) => path.join(inputDir, file));
    console.log(
      chalk.blue(
        `Found ${filePaths.length} PDF files in '${inputDir}' directory.`,
      ),
    );
  } else {
    console.error(
      chalk.red(
        `Input directory '${inputDir}' not found or is not a directory.`,
      ),
    );
    process.exit(1);
  }
} catch (err) {
  console.error(
    chalk.red(`Error reading from input directory '${inputDir}':`),
    err,
  );
  process.exit(1);
}

if (filePaths.length === 0) {
  console.log(chalk.yellow("No PDF files to process."));
  process.exit(0);
}

async function main() {
  if (!fs.existsSync(outputDir)) {
    try {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(chalk.green(`Created output directory: ${outputDir}`));
    } catch (err) {
      console.error(
        chalk.red(`Error creating output directory ${outputDir}:`),
        err,
      );
      process.exit(1);
    }
  }

  const progressBar = new cliProgress.SingleBar({
    format: `${chalk.cyan("{bar}")} | {percentage}% | {value}/{total} | {filename}`,
    barCompleteChar: "\u2588",
    barIncompleteChar: "\u2591",
    hideCursor: true,
  });

  progressBar.start(filePaths.length, 0, {
    filename: "N/A",
  });

  for (const filePath of filePaths) {
    try {
      progressBar.update({ filename: path.basename(filePath) });
      const markdownContent = await convertPdfToMarkdown(filePath);
      const baseName = path.basename(filePath, path.extname(filePath));
      const outputFilePath = path.join(outputDir, `${baseName}.md`);

      fs.writeFileSync(outputFilePath, markdownContent);
    } catch (err) {
      progressBar.stop();
      console.error(chalk.red(`\nError processing ${filePath}:`), err.message);
    } finally {
      progressBar.increment();
    }
  }

  progressBar.stop();
  console.log(chalk.green("\nConversion complete."));
}

main().catch((err) => {
  console.error(chalk.red("An unexpected error occurred:"), err);
  process.exit(1);
});