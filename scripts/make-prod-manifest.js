#!/usr/bin/env node
const fs = require("fs");
const path = require("path");

const DEV_ORIGIN = "https://localhost:3000";
const sourcePath = path.join(__dirname, "..", "manifest.xml");
const outputPath = path.join(__dirname, "..", "manifest.prod.xml");

function fail(message) {
  console.error(`Error: ${message}`);
  console.error(
    "Usage: ADDIN_BASE_URL=https://addin.example.com pnpm manifest:prod"
  );
  process.exit(1);
}

const baseUrl = process.env.ADDIN_BASE_URL;
if (!baseUrl) {
  fail("ADDIN_BASE_URL is not set.");
}

let url;
try {
  url = new URL(baseUrl);
} catch {
  fail(`ADDIN_BASE_URL="${baseUrl}" is not a valid URL.`);
}

if (url.protocol !== "https:" || (url.pathname !== "/" && url.pathname !== "")) {
  fail('ADDIN_BASE_URL must be an https:// origin without a path, e.g. "https://addin.example.com".');
}

const origin = url.origin;
const manifest = fs.readFileSync(sourcePath, "utf8");

if (!manifest.includes(DEV_ORIGIN)) {
  fail(`No occurrences of ${DEV_ORIGIN} found in manifest.xml.`);
}

const count = manifest.split(DEV_ORIGIN).length - 1;
fs.writeFileSync(outputPath, manifest.split(DEV_ORIGIN).join(origin));
console.log(
  `Wrote ${path.basename(outputPath)}: replaced ${count} occurrence(s) of ${DEV_ORIGIN} with ${origin}.`
);
