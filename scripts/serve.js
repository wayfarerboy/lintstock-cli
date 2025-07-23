const express = require("express");
const path = require("node:path");
const livereload = require("livereload");
const connectLiveReload = require("connect-livereload");
const fs = require("node:fs");
const { marked } = require("marked");
const chokidar = require("chokidar");

const app = express();
const port = 3000;
const outputDir = path.join(__dirname, "..", "output");

// Setup live reload
const liveReloadServer = livereload.createServer();
liveReloadServer.server.once("connection", () => {
  setTimeout(() => {
    liveReloadServer.refresh("/");
  }, 100);
});

// Watch for changes in the output directory
const watcher = chokidar.watch(outputDir);
watcher.on("all", (event, path) => {
  liveReloadServer.refresh("/");
});

app.use(connectLiveReload());

// Middleware to render Markdown files
app.use((req, res, next) => {
	const filePath = path.join(outputDir, req.path);
	if (filePath.endsWith(".md") && fs.existsSync(filePath)) {
		fs.readFile(filePath, "utf8", (err, data) => {
			if (err) {
				return next(err);
			}
			const html = marked(data);
			res.send(`
        <!DOCTYPE html>
        <html>
        <head>
          <title>${path.basename(req.path)}</title>
          <style>
            body { font-family: sans-serif; line-height: 1.6; padding: 2em; }
          </style>
        </head>
        <body>
          ${html}
        </body>
        </html>
      `);
		});
	} else {
		next();
	}
});

// Serve static files from the output directory
app.use(express.static(outputDir));

// Directory listing for the root
app.get("/", (req, res) => {
	fs.readdir(outputDir, (err, files) => {
		if (err) {
			res.status(500).send("Error reading output directory");
			return;
		}
		const fileList = files
			.filter((file) => file !== "GEMINI.md")
			.map((file) => `<li><a href="/${file}">${file}</a></li>`)
			.join("");
		res.send(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Output Directory</title>
        <style>
          body { font-family: sans-serif; line-height: 1.6; padding: 2em; }
          ul { list-style-type: none; padding: 0; }
          li { margin: 0.5em 0; }
        </style>
      </head>
      <body>
        <h1>Files in output/</h1>
        <ul>${fileList}</ul>
      </body>
      </html>
    `);
	});
});

app.listen(port, () => {
	console.log(`Serving files from "${outputDir}" on http://localhost:${port}`);
});
