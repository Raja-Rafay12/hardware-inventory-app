const fs = require('fs');

const content = fs.readFileSync('src/HardwareInventoryApp.jsx', 'utf8');

// Find the start of SEED_PRODUCTS
const startMatch = content.indexOf('const SEED_PRODUCTS = [');
if (startMatch === -1) {
  console.log("Could not find SEED_PRODUCTS definition");
  process.exit(1);
}

// Find the matching closing bracket
let openBrackets = 1;
let index = startMatch + 'const SEED_PRODUCTS = ['.length;
let arrayContent = '';

while (openBrackets > 0 && index < content.length) {
  const char = content[index];
  if (char === '[') openBrackets++;
  if (char === ']') openBrackets--;
  arrayContent += char;
  index++;
}

// Evaluate the array content safely as JS
const seedProducts = eval('[' + arrayContent);
console.log("Total seed products in file:", seedProducts.length);
console.log("73rd product (index 72):", seedProducts[72]);

// Print any products that contain characters outside typical ASCII
const nonAsciiProducts = [];
for (let i = 0; i < seedProducts.length; i++) {
  const p = seedProducts[i];
  if (/[^\x00-\x7F]/.test(p.name)) {
    nonAsciiProducts.push({ index: i, id: p.id, name: p.name });
  }
}
console.log(`Found ${nonAsciiProducts.length} products with non-ASCII characters:`);
console.log(nonAsciiProducts.slice(0, 20));
