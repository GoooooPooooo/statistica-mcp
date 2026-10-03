import { inflateSync, deflateSync } from 'node:zlib'

// STATISTICA's graph module exposes no working PDF writer in headless mode
// (SaveAs/SaveAsFormat/SaveAsPDF all return False), so a PDF is produced here by
// embedding the exported PNG as a FlateDecode image on a single page.

const PNG_SIG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])

function decodePng(buf) {
  if (buf.length < 8 || !buf.subarray(0, 8).equals(PNG_SIG)) throw new Error('not a PNG file')
  let pos = 8
  let width = 0
  let height = 0
  let bitDepth = 0
  let colorType = 0
  let interlace = 0
  const idat = []
  while (pos + 8 <= buf.length) {
    const len = buf.readUInt32BE(pos)
    const type = buf.toString('ascii', pos + 4, pos + 8)
    const data = buf.subarray(pos + 8, pos + 8 + len)
    pos += 12 + len
    if (type === 'IHDR') {
      width = data.readUInt32BE(0)
      height = data.readUInt32BE(4)
      bitDepth = data[8]
      colorType = data[9]
      interlace = data[12]
    } else if (type === 'IDAT') {
      idat.push(data)
    } else if (type === 'IEND') {
      break
    }
  }
  if (bitDepth !== 8) throw new Error(`unsupported PNG bit depth ${bitDepth}`)
  if (interlace !== 0) throw new Error('interlaced PNG is not supported')
  const channels = { 0: 1, 2: 3, 4: 2, 6: 4 }[colorType]
  if (!channels) throw new Error(`unsupported PNG color type ${colorType}`)

  const raw = inflateSync(Buffer.concat(idat))
  const stride = width * channels
  const img = Buffer.alloc(height * stride)
  let rp = 0
  for (let y = 0; y < height; y++) {
    const filter = raw[rp++]
    const rowStart = y * stride
    const prevStart = rowStart - stride
    for (let x = 0; x < stride; x++) {
      const cur = raw[rp + x]
      const a = x >= channels ? img[rowStart + x - channels] : 0
      const b = y > 0 ? img[prevStart + x] : 0
      const c = y > 0 && x >= channels ? img[prevStart + x - channels] : 0
      let val
      switch (filter) {
        case 0:
          val = cur
          break
        case 1:
          val = cur + a
          break
        case 2:
          val = cur + b
          break
        case 3:
          val = cur + ((a + b) >> 1)
          break
        case 4: {
          const p = a + b - c
          const pa = Math.abs(p - a)
          const pb = Math.abs(p - b)
          const pc = Math.abs(p - c)
          val = cur + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c)
          break
        }
        default:
          throw new Error(`unknown PNG filter ${filter}`)
      }
      img[rowStart + x] = val & 0xff
    }
    rp += stride
  }

  // Flatten to 8-bit RGB (drop any alpha channel).
  const rgb = Buffer.alloc(width * height * 3)
  for (let i = 0, o = 0; i < width * height; i++, o += 3) {
    const s = i * channels
    if (channels === 1) {
      rgb[o] = rgb[o + 1] = rgb[o + 2] = img[s]
    } else if (channels === 2) {
      rgb[o] = rgb[o + 1] = rgb[o + 2] = img[s]
    } else {
      rgb[o] = img[s]
      rgb[o + 1] = img[s + 1]
      rgb[o + 2] = img[s + 2]
    }
  }
  return { width, height, rgb }
}

export function pngToPdf(png) {
  const { width, height, rgb } = decodePng(png)
  const z = deflateSync(rgb)

  const objects = []
  objects[1] = Buffer.from('<< /Type /Catalog /Pages 2 0 R >>')
  objects[2] = Buffer.from('<< /Type /Pages /Kids [3 0 R] /Count 1 >>')
  objects[3] = Buffer.from(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${width} ${height}] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>`)
  objects[4] = Buffer.concat([
    Buffer.from(`<< /Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length ${z.length} >>\nstream\n`),
    z,
    Buffer.from('\nendstream'),
  ])
  const content = `q\n${width} 0 0 ${height} 0 0 cm\n/Im0 Do\nQ\n`
  objects[5] = Buffer.from(`<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}endstream`)

  const chunks = [Buffer.from([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x34, 0x0a, 0x25, 0xff, 0xff, 0xff, 0xff, 0x0a])]
  let length = chunks[0].length
  const offsets = [0]
  for (let i = 1; i <= 5; i++) {
    offsets[i] = length
    const obj = Buffer.concat([Buffer.from(`${i} 0 obj\n`), objects[i], Buffer.from('\nendobj\n')])
    chunks.push(obj)
    length += obj.length
  }
  const xrefStart = length
  let xref = 'xref\n0 6\n0000000000 65535 f \n'
  for (let i = 1; i <= 5; i++) xref += `${String(offsets[i]).padStart(10, '0')} 00000 n \n`
  xref += `trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF\n`
  chunks.push(Buffer.from(xref))

  return Buffer.concat(chunks)
}
