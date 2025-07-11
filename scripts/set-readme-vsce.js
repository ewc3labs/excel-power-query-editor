
/**
 * Copies the README.vsmarketplace.md file from the docs directory
 * to the root directory as README.md. This is typically used to
 * set the README file for publishing to the VS Marketplace.
 *
 * Source: ../docs/README.vsmarketplace.md
 * Destination: ../README.md
 *
 * Logs a message upon successful copy.
 */
const fs = require('fs');
const path = require('path');

function setReadmeVSMarketplace() {
  const source = path.join(__dirname, '..', 'docs', 'README.vsmarketplace.md');
  const dest = path.join(__dirname, '..', 'README.md');
  fs.copyFileSync(source, dest);
  console.log('[set-readme-vsmarketplace] Set VS Marketplace README');
}

setReadmeVSMarketplace();
