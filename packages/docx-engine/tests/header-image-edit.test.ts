import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'

/** edits of images already in a header/footer part (imageEdits): the fix for
 *  image-only logo headers that could display but never be removed/replaced */

const HEADER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

/** header shaped like a real Word logo header: one centered paragraph holding one inline drawing */
async function logoHeaderDoc() {
  const drawing =
    '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
    '<wp:extent cx="1905000" cy="244475"/>' +
    '<wp:docPr id="1" name="Logo"/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:nvPicPr><pic:cNvPr id="1" name="Logo"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="rIdHF"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="244475"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  const HEADER =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    `<w:p><w:pPr><w:spacing w:after="200"/><w:jc w:val="center"/></w:pPr>${drawing}</w:p>` +
    '</w:hdr>'
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    extraRels: `<Relationship Id="rId61" Type="${HEADER_REL}" Target="header1.xml"/>`,
    extraParts: [
      {
        path: 'word/header1.xml',
        xml: HEADER,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
      },
      {
        path: 'word/_rels/header1.xml.rels',
        xml:
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="rIdHF" Type="${IMAGE_REL}" Target="media/logo.png"/>` +
          '</Relationships>',
        contentType: '',
      },
    ],
    binaryParts: [
      {
        path: 'word/media/logo.png',
        base64: TINY_PNG_BASE64,
        extension: 'png',
        contentType: 'image/png',
      },
    ],
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId61"/>',
  })
  const parsed = await parseDocx(bytes)
  const saveBlocks: SaveBlock[] = [{ kind: 'original', docxIndex: parsed.blocks[0].docxIndex! }]
  return { parsed, saveBlocks }
}

const headerXmlOf = async (docx: Uint8Array): Promise<string> => {
  const zip = await JSZip.loadAsync(docx)
  return (await zip.file('word/header1.xml')?.async('string')) ?? ''
}

describe('header imageEdits', () => {
  it('removes an image-only header paragraph (the logo header case)', async () => {
    const { parsed, saveBlocks } = await logoHeaderDoc()
    expect(parsed.headerImages).toHaveLength(1)
    const saved = await saveDocx(parsed, saveBlocks, {
      header: { text: '', imageEdits: [{ index: 0, op: 'remove' }] },
    })
    const reparsed = await parseDocx(saved)
    expect(reparsed.headerImages ?? []).toHaveLength(0)
    const xml = await headerXmlOf(saved)
    expect(xml).not.toContain('<w:drawing')
    // the emptied centered paragraph does not linger as a blank header line
    expect(xml).not.toMatch(/<w:p><w:pPr><w:spacing[^/]*\/><w:jc w:val="center"\/><\/w:pPr><\/w:p>/)
  })

  it('replaces the drawing run and keeps the paragraph formatting (jc=center)', async () => {
    const { parsed, saveBlocks } = await logoHeaderDoc()
    const saved = await saveDocx(parsed, saveBlocks, {
      header: {
        text: '',
        imageEdits: [
          {
            index: 0,
            op: 'replace',
            image: { base64: TINY_PNG_BASE64, mime: 'image/png', widthPx: 80, heightPx: 20 },
          },
        ],
      },
    })
    const reparsed = await parseDocx(saved)
    expect(reparsed.headerImages ?? []).toHaveLength(1)
    const xml = await headerXmlOf(saved)
    // the paragraph keeps its original centering
    expect(xml).toContain('<w:jc w:val="center"/>')
    // the old media reference is gone; a fresh relationship points at new media
    expect(xml).not.toContain('r:embed="rIdHF"')
    expect(xml).toMatch(/<a:blip r:embed="[^"]+"\/>/)
    // the stored extent is the new image's size (80x20px -> EMU)
    expect(xml).toContain(`cx="${80 * 9525}"`)
    // the part's own rels gained the new image relationship
    const zip = await JSZip.loadAsync(saved)
    const rels = (await zip.file('word/_rels/header1.xml.rels')?.async('string')) ?? ''
    expect(rels).toMatch(/Target="media\/[^"]+"/)
    // a new media part landed in the zip
    expect(
      Object.keys(zip.files).some((n) =>
        /^word\/media\/(aidocs|image)\d+\.(png|jpeg|gif)$/.test(n),
      ),
    ).toBe(true)
  })

  it('without imageEdits the image paragraph keeps its original bytes', async () => {
    const { parsed, saveBlocks } = await logoHeaderDoc()
    const saved = await saveDocx(parsed, saveBlocks, { header: { text: '' } })
    const xml = await headerXmlOf(saved)
    expect(xml).toContain('r:embed="rIdHF"')
    expect(xml).toContain('<w:jc w:val="center"/>')
  })
})
