'use strict';

const fs = require('fs');
const path = require('path');

const source = path.join(__dirname, '..', 'temp', 'build', 'manifests.js');
const destination = path.join(__dirname, '..', 'temp', 'manifests.js');
const bundlePath = path.join(__dirname, '..', 'dist', 'mi360-data-security-gate-application-customizer_1ab5f7de57a2e5c01c82.js');
const bundleName = 'mi360-data-security-gate-application-customizer';
const componentModule = 'f0f108c2-8e86-42f5-993d-cb1863c8f6d4_1.0.0.0';

if (!fs.existsSync(source)) {
  throw new Error(`Debug manifest was not generated: ${source}`);
}

let manifest = fs.readFileSync(source, 'utf8');

if (fs.existsSync(bundlePath)) {
  let bundle = fs.readFileSync(bundlePath, 'utf8');
  const aliasDefinition = `define("${bundleName}"`;
  if (!bundle.includes(aliasDefinition)) {
    bundle += `\ndefine("${bundleName}",["${componentModule}"],function(component){return component;});\n`;
    fs.writeFileSync(bundlePath, bundle);
  }
}

fs.writeFileSync(destination, manifest);
console.log(`Copied ${source} to ${destination}`);
