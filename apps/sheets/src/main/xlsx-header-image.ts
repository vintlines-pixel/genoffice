/// Reads a worksheet's image header (&C&G + <legacyDrawingHF>) out of an
/// xlsx package: worksheet part → VML drawing → media bytes. Excel stores
/// picture headers as a VML shape whose imagedata points at a media part
/// through the VML drawing's own rels. Returns null when the sheet has no
/// picture header (text-only headers never come through here).
import { readFileSync } from 'node:fs'
import JSZip from 'jszip'

export interface SheetHeaderImage {
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
  base64: string
}

const XML_ATTR = (tag: string, attr: string): RegExp =>
  new RegExp(`<${tag}\\b[^>]*\\b${attr}="([^"]*)"[^>]*>`)

const decodeXmlEntities = (value: string): string =>
  value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')

const relTarget = (relsXml: string, relId: string): string | null =>
  new RegExp(`<Relationship\\b[^>]*\\bId="${relId}"[^>]*\\bTarget="([^"]*)"`).exec(relsXml)?.[1] ??
  null

const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/emf',
  wmf: 'image/wmf',
}

/** Picture header of one sheet (by its workbook-visible name); null when absent. */
export async function readSheetHeaderImage(
  filePath: string,
  sheetName: string,
): Promise<SheetHeaderImage | null> {
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(readFileSync(filePath))
  } catch {
    return null
  }
  try {
    const workbookXml = await zip.file('xl/workbook.xml')?.async('string')
    const workbookRels = await zip.file('xl/_rels/workbook.xml.rels')?.async('string')
    if (!workbookXml || !workbookRels) return null
    // resolve the sheet name to its worksheet part
    for (const m of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
      const tag = m[0]
      const name = /name="([^"]*)"/.exec(tag)?.[1]
      if (name === undefined || decodeXmlEntities(name) !== sheetName) continue
      const rid = /r:id="([^"]*)"/.exec(tag)?.[1]
      if (!rid) continue
      const target = relTarget(workbookRels, rid)
      if (!target) continue
      const partPath = `xl/${target.replace(/^\//, '').replace(/^xl\//, '')}`
      const sheetXml = await zip.file(partPath)?.async('string')
      if (!sheetXml) return null
      // <legacyDrawingHF r:id="rId2"/> → the VML drawing holding header shapes
      const legacy = XML_ATTR('legacyDrawingHF', 'r:id').exec(sheetXml)?.[1]
      if (!legacy) return null
      const sheetRelsPath = partPath.replace(/([^/]+)$/, '_rels/$1.rels')
      const sheetRels = await zip.file(sheetRelsPath)?.async('string')
      if (!sheetRels) return null
      const vmlTarget = relTarget(sheetRels, legacy)
      if (!vmlTarget) return null
      const vmlPath = `xl/${vmlTarget
        .replace(/^\//, '')
        .replace(/^\.\.\//, '')
        .replace(/^xl\//, '')}`
      const vmlXml = await zip.file(vmlPath)?.async('string')
      if (!vmlXml || !vmlXml.includes('imagedata')) return null
      const vmlRelsPath = vmlPath.replace(/([^/]+)$/, '_rels/$1.rels')
      const vmlRels = await zip.file(vmlRelsPath)?.async('string')
      if (!vmlRels) return null
      // VML references the image via o:relid or r:id
      const imgRel =
        XML_ATTR('v:imagedata', 'o:relid').exec(vmlXml)?.[1] ??
        XML_ATTR('v:imagedata', 'r:id').exec(vmlXml)?.[1]
      if (!imgRel) return null
      const imgTarget = relTarget(vmlRels, imgRel)
      if (!imgTarget) return null
      const mediaPath = `xl/${imgTarget
        .replace(/^\//, '')
        .replace(/^(\.\.\/)+/, '')
        .replace(/^xl\//, '')}`
      const file = zip.file(mediaPath)
      if (!file) return null
      const ext = (mediaPath.split('.').pop() ?? '').toLowerCase()
      const mime = MIME_BY_EXT[ext]
      // only persistable formats render in the print header
      if (mime !== 'image/png' && mime !== 'image/jpeg' && mime !== 'image/gif') return null
      const bytes = await file.async('base64')
      if (!bytes) return null
      return { mime: mime as SheetHeaderImage['mime'], base64: bytes }
    }
    return null
  } catch {
    return null
  }
}
