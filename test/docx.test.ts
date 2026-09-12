import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { extractDocxImages, resolveTarget } from "../src/core/docx";
import { nodePlatform } from "../src/node/platform-node";
import { editDocumentXml, SAMPLE_DOCX } from "./helpers";

describe("extractDocxImages", () => {
  it("finds the embedded picture, its size, display extent and name", async () => {
    const { parts, usages } = await extractDocxImages(SAMPLE_DOCX, nodePlatform);
    assert.equal(parts.length, 1);
    assert.equal(parts[0].path, "word/media/image1.jpeg");
    assert.equal(parts[0].mime, "image/jpeg");
    assert.equal(parts[0].width, 2800);
    assert.equal(parts[0].height, 1867);

    assert.equal(usages.length, 1);
    const [u] = usages;
    assert.equal(u.storyPart, "word/document.xml");
    assert.equal(u.name, "Picture");
    assert.deepEqual(u.extentEmu, { cx: 6120130, cy: 4080510 });
    assert.equal(u.crop, undefined);
    assert.deepEqual(u.blockers, []);
  });

  it("reads a:srcRect crops and flags effects", async () => {
    const docx = editDocumentXml(SAMPLE_DOCX, (xml) =>
      xml.replace(
        '<a:blip r:embed="rId4"/>',
        '<a:blip r:embed="rId4"><a:lum bright="20000"/></a:blip><a:srcRect l="5000" r="5000"/>',
      ),
    );
    const { usages } = await extractDocxImages(docx, nodePlatform);
    assert.deepEqual(usages[0].crop, { left: 0.05, top: 0, right: 0.05, bottom: 0 });
    assert.deepEqual(usages[0].blockers, ["picture effect (lum)"]);
  });

  it("ignores harmless a:extLst entries such as a14:useLocalDpi", async () => {
    const docx = editDocumentXml(SAMPLE_DOCX, (xml) =>
      xml.replace(
        '<a:blip r:embed="rId4"/>',
        '<a:blip r:embed="rId4"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip>',
      ),
    );
    const { usages } = await extractDocxImages(docx, nodePlatform);
    assert.deepEqual(usages[0].blockers, []);
  });

  it("resolves relationship targets relative to the source part", () => {
    assert.equal(resolveTarget("word/document.xml", "media/image1.png"), "word/media/image1.png");
    assert.equal(resolveTarget("word/document.xml", "../media/x.png"), "media/x.png");
    assert.equal(resolveTarget("word/document.xml", "/word/media/y.png"), "word/media/y.png");
    assert.equal(resolveTarget("", "word/document.xml"), "word/document.xml");
  });
});
