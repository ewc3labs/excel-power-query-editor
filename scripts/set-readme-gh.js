/**
 * The 'fs' module provides an API for interacting with the file system.
 * It allows reading, writing, updating, and deleting files and directories.
 * 
 * @module fs
 * @see {@link https://nodejs.org/api/fs.html|Node.js fs documentation}
 */
const fs = require('fs');
const path = require('path');

function setReadmeGH() {
  const source = path.join(__dirname, '..', 'docs', 'README.gh.md');
  const dest = path.join(__dirname, '..', 'README.md');
  fs.copyFileSync(source, dest);
  console.log('[set-readme-gh] Restored GitHub README');
}

setReadmeGH();
