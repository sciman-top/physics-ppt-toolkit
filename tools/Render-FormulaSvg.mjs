#!/usr/bin/env node

import fs from 'node:fs/promises';
import path from 'node:path';

function parseArgs(argv) {
  const args = {
    tex: '',
    out: '',
    display: true,
    em: 32,
    ex: 16,
    widthEm: 80,
  };

  for (let i = 2; i < argv.length; i++) {
    const key = argv[i];
    const next = argv[i + 1];
    if (key === '--tex') {
      args.tex = next || '';
      i++;
    } else if (key === '--out') {
      args.out = next || '';
      i++;
    } else if (key === '--inline') {
      args.display = false;
    } else if (key === '--em') {
      args.em = Number(next || args.em);
      i++;
    } else if (key === '--width-em') {
      args.widthEm = Number(next || args.widthEm);
      i++;
    } else if (key === '--help' || key === '-h') {
      args.help = true;
    } else {
      throw new Error(`Unknown argument: ${key}`);
    }
  }

  return args;
}

function printUsage() {
  process.stdout.write(`Usage:
  node tools/Render-FormulaSvg.mjs --tex "P=\\\\frac{W}{t}" --out reports/formula.svg

Options:
  --tex       TeX formula source.
  --out       Output SVG path.
  --inline    Render in inline mode instead of display mode.
  --em        MathJax em size in px. Default: 32.
  --width-em  Container width in em. Default: 80.
`);
}

async function ensureOutputDir(outputPath) {
  const dir = path.dirname(path.resolve(outputPath));
  await fs.mkdir(dir, { recursive: true });
}

globalThis.MathJax = {
  loader: {
    paths: { mathjax: '@mathjax/src/bundle' },
    load: ['adaptors/liteDOM'],
    require: (file) => import(file),
  },
  options: {
    enableSpeech: false,
    enableBraille: false,
  },
  output: {
    font: 'mathjax-newcm',
  },
};

await import('@mathjax/src/bundle/tex-svg.js');
await globalThis.MathJax.startup.promise;

async function texToSvg(tex, options) {
  const node = await globalThis.MathJax.tex2svgPromise(tex, {
    display: options.display,
    em: options.em,
    ex: options.ex,
    containerWidth: options.widthEm * options.em,
  });
  const adaptor = globalThis.MathJax.startup.adaptor;
  const svgNode = adaptor.tags(node, 'svg')[0];
  let svg = adaptor.serializeXML(svgNode);
  if (!svg.includes('xmlns=')) {
    svg = svg.replace('<svg ', '<svg xmlns="http://www.w3.org/2000/svg" ');
  }
  return `${svg}\n`;
}

try {
  const args = parseArgs(process.argv);
  if (args.help) {
    printUsage();
    process.exit(0);
  }
  if (!args.tex.trim()) {
    throw new Error('Missing --tex.');
  }
  if (!args.out.trim()) {
    throw new Error('Missing --out.');
  }

  const svg = await texToSvg(args.tex, args);
  await ensureOutputDir(args.out);
  await fs.writeFile(args.out, svg, 'utf8');
  process.stdout.write(`SVG written: ${path.resolve(args.out)}\n`);
} finally {
  globalThis.MathJax.done();
}
