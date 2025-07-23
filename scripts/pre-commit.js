const { execSync } = require("node:child_process");

try {
  console.log("Bumping version number...");
  execSync("npm version patch --no-git-tag-version");
  execSync("git add package.json");
  console.log("Version bumped and package.json staged.");
} catch (error) {
  console.error("Failed to bump version:", error);
  process.exit(1);
}
