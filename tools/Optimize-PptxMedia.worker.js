const fs = require('fs');
const sharp = require('sharp');

async function main() {
  const argsPath = process.argv[2];
  if (!argsPath) {
    throw new Error('Missing args path.');
  }

  const raw = fs.readFileSync(argsPath, 'utf8').replace(/^\uFEFF/, '');
  const args = JSON.parse(raw);

  let pipeline = sharp(args.input, { failOn: 'none' }).rotate();
  if (args.sharpen) {
    pipeline = pipeline.sharpen();
  }

  if (args.kind === 'jpeg') {
    await pipeline.jpeg({ quality: args.quality, mozjpeg: true }).toFile(args.output);
  } else if (args.kind === 'png') {
    await pipeline.png({ compressionLevel: 9, adaptiveFiltering: true }).toFile(args.output);
  } else {
    throw new Error(`Unsupported media kind: ${args.kind}`);
  }

  const metadata = await sharp(args.output, { failOn: 'none' }).metadata();
  process.stdout.write(JSON.stringify({
    width: metadata.width || 0,
    height: metadata.height || 0,
    format: metadata.format || ''
  }));
}

main().catch((error) => {
  process.stderr.write(error && error.stack ? error.stack : String(error));
  process.exit(1);
});
