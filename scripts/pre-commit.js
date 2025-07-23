const { execSync } = require("node:child_process");

try {
  console.log("Bumping version number...");
  execSync("npm version patch --no-git-tag-version");
  execSync("git add package.json package-lock.json");
  console.log("Version bumped and package.json and package-lock.json staged.");
} catch (error) {
  console.error("Failed to bump version:", error);
  process.exit(1);
}
