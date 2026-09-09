const fs = require('fs');
const sharp = require('sharp');

async function main() {
  const argsPath = process.argv[2];
  if (!argsPath) {
    throw new Error('Missing args path.');
  }

  const raw = fs.readFileSync(argsPath, 'utf8').replace(/^\uFEFF/, '');
  const args = JSON.parse(raw);

  // Preserves EXIF orientation: pixels stay as encoded and the orientation tag
  // is kept, so viewers render the optimized copy exactly like the original.
  // .rotate() would auto-orient pixels and swap width/height instead.
  let pipeline = sharp(args.input, { failOn: 'none' }).withMetadata();
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
  // exitCode instead of exit(1) so pending stderr writes drain before exit.
  process.exitCode = 1;
});
