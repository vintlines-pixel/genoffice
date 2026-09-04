/**
 * Inserted-image downscale: a document/worksheet holding a multi-megapixel
 * original (e.g. a full-screen screenshot) re-decodes and re-lays it on every
 * edit, which freezes the editor. Insert-time re-encoding to a bounded size
 * keeps display/print quality while cutting the decode + layout + save cost.
 *
 * Policy: PNG stays lossless PNG, JPEG re-encodes as JPEG, animated GIF is
 * left untouched (a single frame would kill the animation). Tiny images are
 * returned as-is. Any decode/canvas failure falls back to the original.
 */

/** Long-edge cap for an inserted image (≈ print a 12cm wide picture at 270 dpi). */
export const MAX_INSERT_IMAGE_DIM = 1280
export const INSERT_IMAGE_JPEG_QUALITY = 0.85

export interface FittedImage {
  width: number
  height: number
}

/** Scale so the long edge lands inside maxDim; never upscales. */
export function fitImageDim(width: number, height: number, maxDim: number): FittedImage {
  const long = Math.max(width, height)
  if (long <= 0 || long <= maxDim) return { width, height }
  const scale = maxDim / long
  return {
    width: Math.max(1, Math.round(width * scale)),
    height: Math.max(1, Math.round(height * scale)),
  }
}

export interface DownscaledImage {
  dataUrl: string
  width: number
  height: number
}

function decodeSize(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image()
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight })
    img.onerror = () => resolve({ width: 0, height: 0 })
    img.src = dataUrl
  })
}

function reencode(
  dataUrl: string,
  width: number,
  height: number,
  mime: 'image/png' | 'image/jpeg',
  quality: number,
): Promise<DownscaledImage> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => {
      try {
        const canvas = document.createElement('canvas')
        canvas.width = width
        canvas.height = height
        const ctx = canvas.getContext('2d')
        if (!ctx) throw new Error('no 2d context')
        ctx.drawImage(img, 0, 0, width, height)
        resolve({
          dataUrl: canvas.toDataURL(mime, quality),
          width,
          height,
        })
      } catch (err) {
        reject(err)
      }
    }
    img.onerror = () => reject(new Error('decode failed'))
    img.src = dataUrl
  })
}

/** Downscale a data-URL image to MAX_INSERT_IMAGE_DIM on the long edge. */
export async function downscaleImageDataUrl(
  dataUrl: string,
  opts?: { maxDim?: number; quality?: number },
): Promise<DownscaledImage> {
  const match = /^data:(image\/(?:png|jpeg|gif));/i.exec(dataUrl)
  const mime = (match?.[1] ?? '').toLowerCase() as 'image/png' | 'image/jpeg' | 'image/gif' | ''
  if (mime === '') return { dataUrl, width: 0, height: 0 }
  const maxDim = opts?.maxDim ?? MAX_INSERT_IMAGE_DIM
  const natural = await decodeSize(dataUrl)
  if (!natural.width || !natural.height) return { dataUrl, width: 0, height: 0 }
  const { width, height } = fitImageDim(natural.width, natural.height, maxDim)
  if (mime === 'image/gif') return { dataUrl, width: natural.width, height: natural.height }
  if (width === natural.width && height === natural.height) {
    return { dataUrl, width: natural.width, height: natural.height }
  }
  try {
    return await reencode(
      dataUrl,
      width,
      height,
      mime,
      opts?.quality ?? INSERT_IMAGE_JPEG_QUALITY,
    )
  } catch {
    return { dataUrl, width: natural.width, height: natural.height }
  }
}
