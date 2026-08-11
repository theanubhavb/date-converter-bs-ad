const fs = require("fs");
const path = require("path");

const [owner, repository = "date-converter-bs-ad"] = process.argv.slice(2);
if (!owner || !/^[A-Za-z0-9-]+$/.test(owner) || !/^[A-Za-z0-9._-]+$/.test(repository)) {
  throw new Error("Usage: npm run configure:github-pages -- <github-owner> [repository-name]");
}

const root = path.resolve(__dirname, "..");
const host = `https://${owner}.github.io/${repository}`;
const source = fs.readFileSync(path.join(root, "manifest.xml"), "utf8");
const manifest = source
  .replaceAll("https://localhost:3000", host)
  .replaceAll("YOUR_GITHUB_ACCOUNT/date_converter_bs_ad", `${owner}/${repository}`);
fs.writeFileSync(path.join(root, "manifest.github-pages.xml"), manifest);
console.log(`Created manifest.github-pages.xml for ${host}`);
