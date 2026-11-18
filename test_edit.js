const fs = require('fs');
const content = fs.readFileSync('app.js', 'utf8');
const lines = content.split('\n');
for (let i = 1290; i < 1300; i++) {
  console.log((i+1).toString().padStart(4), JSON.stringify(lines[i]));
}