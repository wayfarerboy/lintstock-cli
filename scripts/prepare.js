const fs = require("node:fs");
const path = require("node:path");
const chalk = require("chalk");
const cliProgress = require("cli-progress");
const { parseSpreadsheetBuffer } = require("./spreadsheet-to-json");

// Example (single file, filtered): node scripts/prepare.js --question 4 --outdir ./output data/some-file.xlsx
// Example (multiple files, unfiltered): node scripts/prepare.js data/file1.xlsx data/file2.xlsx
// Example (all files in data/, unfiltered): node scripts/prepare.js

async function processLocalExcelFile(localFilePath) {
	const fileContent = fs.readFileSync(localFilePath);
	return parseSpreadsheetBuffer(fileContent);
}

// --- Argument Parsing ---
const args = process.argv.slice(2);
let filePaths = [];
let questionNumber = null;
let outDir = "context/reports"; // Default output directory

for (let i = 0; i < args.length; i++) {
	if (args[i] === "--question") {
		const qNum = args[i + 1];
		if (qNum && !qNum.startsWith("-")) {
			const parsed = Number.parseInt(qNum, 10);
			if (!Number.isNaN(parsed)) {
				questionNumber = parsed;
			}
			i++; // Skip the value
		}
	} else if (args[i] === "--outdir") {
		const dir = args[i + 1];
		if (dir && !dir.startsWith("-")) {
			outDir = dir;
			i++; // Skip the value
		}
	} else {
		filePaths.push(args[i]);
	}
}

// If no file paths are provided, default to all files in the 'data' directory
if (filePaths.length === 0) {
	const dataDir = "data/reports";
	try {
		if (fs.existsSync(dataDir) && fs.statSync(dataDir).isDirectory()) {
			const allFiles = fs.readdirSync(dataDir);
			// Filter out subdirectories and non-spreadsheet files if necessary
			filePaths = allFiles
				.map((file) => path.join(dataDir, file))
				.filter((filePath) => {
					const stat = fs.statSync(filePath);
					// Add any other relevant extensions if needed
					return (
						stat.isFile() &&
						(filePath.endsWith(".xlsx") ||
							filePath.endsWith(".xls") ||
							filePath.endsWith(".csv"))
					);
				});
			console.log(
				chalk.blue(
					`No input files specified, processing all compatible files in '${dataDir}' directory.`,
				),
			);
		} else {
			console.error(
				chalk.red(
					`Default data directory '${dataDir}' not found or is not a directory.`,
				),
			);
			process.exit(1);
		}
	} catch (err) {
		console.error(
			chalk.red(`Error reading from data directory '${dataDir}':`),
			err,
		);
		process.exit(1);
	}
}

if (filePaths.length === 0) {
	console.log(
		chalk.yellow(
			"No files to process. Please provide file paths or ensure the 'data/reports' directory has compatible files.",
		),
	);
	process.exit(0);
}

async function main() {
	// Create output directory if it doesn't exist
	if (!fs.existsSync(outDir)) {
		try {
			fs.mkdirSync(outDir, { recursive: true });
			console.log(chalk.green(`Created output directory: ${outDir}`));
		} catch (err) {
			console.error(
				chalk.red(`Error creating output directory ${outDir}:`),
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

	const allUnmatchedHeaders = new Map();

	for (const filePath of filePaths) {
		try {
			progressBar.update({ filename: path.basename(filePath) });
			if (fs.statSync(filePath).isDirectory()) {
				progressBar.increment();
				continue;
			}
			const json = await processLocalExcelFile(filePath);

			if (questionNumber !== null) {
				if (json.reports && Array.isArray(json.reports)) {
					json.reports = json.reports
						.map((report) => ({
							...report,
							questions: report.questions.filter(
								(q) => q.question_number === questionNumber,
							),
						}))
						.filter((report) => report.questions.length > 0);
				}
			}

			const outputJson = JSON.stringify(json, null, 2);
			const baseName = path.basename(filePath, path.extname(filePath));
			const outputFilePath = path.join(outDir, `${baseName}.json`);

			try {
				fs.writeFileSync(outputFilePath, outputJson);
			} catch (writeErr) {
				progressBar.stop();
				console.error(
					chalk.red(`\nError writing to ${outputFilePath}:`),
					writeErr,
				);
			}

			if (json.unmatchedHeaders && json.unmatchedHeaders.length > 0) {
				allUnmatchedHeaders.set(filePath, json.unmatchedHeaders);
			}
		} catch (err) {
			progressBar.stop();
			console.error(chalk.red(`\nError processing ${filePath}:`), err.message);
		} finally {
			progressBar.increment();
		}
	}

	progressBar.stop();

	if (allUnmatchedHeaders.size > 0) {
		console.log(
			chalk.yellow(
				"\n[UNMATCHED HEADERS] The following headers were not mapped:",
			),
		);
		for (const [filePath, headers] of allUnmatchedHeaders.entries()) {
			console.log(chalk.yellow(`\nFile: ${filePath}`));
			for (const h of headers) {
				console.log(chalk.yellow("-", h));
			}
		}
	}
	console.log();
}

main().catch((err) => {
	console.error(chalk.red("An unexpected error occurred:"), err);
	process.exit(1);
});
