import JSZip from 'jszip';
import { parseStringPromise } from 'xml2js';

interface ParagraphItem {
  text: TextItem[];
  isList: boolean;
}

interface TextItem {
  text: string;
  isBold: boolean;
}

interface SlideData {
  title: string;
  text: ParagraphItem[];
  images: string[];
  number: number;
}

interface PluginMessage {
  type: string;
  data: ArrayBuffer;
  name: string;
  posX: number;
}

figma.showUI(__html__);

figma.ui.onmessage = async (msg: PluginMessage) => {
  if (msg.type === 'file') {
    await handleFileMessage(msg);
  }
};

async function handleFileMessage(msg: PluginMessage) {
  if (!(msg.data instanceof ArrayBuffer)) {
    console.error('Received data is not an ArrayBuffer.');
    return;
  }

  try {
    const { zip, name, posX } = await loadZipFile(msg);
    await getSlideGrids();
    const slides = await parseSlides(zip);
    await renderSlidesToFrames(slides, zip, name, posX);
  } catch (error) {
    console.error('An error occurred in the file processing:', error);
  }
}

async function loadZipFile(msg: PluginMessage) {
  try {
    const zip = await JSZip.loadAsync(msg.data);
    return { zip, name: msg.name, posX: msg.posX };
  } catch (err) {
    console.error('Error loading zip file:', err);
    throw err;
  }
}

async function parseSlides(zip: JSZip): Promise<SlideData[]> {
  const slides: SlideData[] = [];

  const slidePaths = getSlidePaths(zip).sort((a, b) => {
    const numA = parseInt(a.match(/slide(\d+)/)?.[1] || '0');
    const numB = parseInt(b.match(/slide(\d+)/)?.[1] || '0');
    return numA - numB;
  });
  const slideToRelsMap = getSlideToRelsMap(zip);

  for (const path of slidePaths) {
    const slideData = await processSlide(zip, path, slideToRelsMap[path]);
    if (slideData) slides.push(slideData);
  }
  return slides;
}

function getSlidePaths(zip: JSZip): string[] {
  return Object.keys(zip.files).filter((path) =>
    path.startsWith('ppt/slides/slide')
  );
}

function getSlideToRelsMap(zip: JSZip): { [key: string]: string } {
  const relsPaths = Object.keys(zip.files).filter((path) =>
    path.startsWith('ppt/slides/_rels/slide')
  );

  const slideToRelsMap: { [key: string]: string } = {};
  for (const relsPath of relsPaths) {
    const slideNumber = relsPath.match(/slide(\d+)/)?.[1];
    if (slideNumber) {
      slideToRelsMap[`ppt/slides/slide${slideNumber}.xml`] = relsPath;
    }
  }
  return slideToRelsMap;
}

async function processSlide(
  zip: JSZip,
  path: string,
  relsPath: string
): Promise<SlideData | null> {
  try {
    const slideFile = zip.file(path);
    const relsFile = zip.file(relsPath);

    if (!slideFile || !relsFile) {
      console.warn(`Slide or rels file not found for path: ${path}`);
      return null;
    }

    const slideContent = await slideFile.async('text');
    const slideData = await parseStringPromise(slideContent);

    const relsContent = await relsFile.async('text');
    const relsData = await parseStringPromise(relsContent);

    const { title, paragraphs } = extractTextData(slideData);

    const images = extractImagePaths(relsData);

    return {
      title,
      text: formatParagraphItems(paragraphs),
      images,
      number: 0,
    };
  } catch (error) {
    console.error(`Error processing slide ${path}:`, error);
    return null;
  }
}

function extractTextData(slideData: any): {
  title: string;
  paragraphs: ParagraphItem[];
} {
  let title = '';
  const paragraphs: ParagraphItem[] = [];

  const shapes =
    slideData?.['p:sld']?.['p:cSld']?.[0]?.['p:spTree']?.[0]?.['p:sp'] || [];
  title =
    shapes.find(
      (shape: any) =>
        shape?.['p:nvSpPr']?.[0]?.['p:nvPr']?.[0]?.[
          'p:ph'
        ]?.[0]?.$?.type.includes('title') ||
        shape?.['p:nvSpPr']?.[0]?.['p:nvPr']?.[0]?.[
          'p:ph'
        ]?.[0]?.$?.type.includes('ctrTitle')
    )?.['p:txBody']?.[0]?.['a:p']?.[0]?.['a:r']?.[0]?.['a:t']?.[0] || '';
  for (const shape of shapes) {
    const textBody = shape?.['p:txBody'];
    if (textBody) {
      const paragraphsInShape = textBody[0]?.['a:p'];
      if (paragraphsInShape) {
        for (const paragraph of paragraphsInShape) {
          const textArray: TextItem[] = [];

          for (const textItem of paragraph?.['a:r'] || []) {
            const text = textItem['a:t']?.[0];
            if (text !== title) {
              textArray.push({
                text,
                isBold: textItem['a:rPr']?.[0]?.$?.b === '1',
              });
            }
          }

          if (!textArray.length) continue;

          const paragraphProps = paragraph?.['a:pPr']?.[0];
          const isList = !paragraphProps || !paragraphProps['a:buNone'];

          paragraphs.push({ text: textArray, isList });
        }
      }
    }
  }

  return { title, paragraphs };
}

function extractImagePaths(relsData: any): string[] {
  const images: string[] = [];
  const relationships = relsData?.['Relationships']?.['Relationship'] || [];
  for (const relationship of relationships) {
    if (
      relationship.$.Type ===
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
    ) {
      const imagePath = `ppt/${relationship.$.Target.replace('../', '')}`;
      if (imagePath) {
        images.push(`${imagePath}`);
      } else {
        console.warn(`Image file not found for path: ${imagePath}`);
      }
    }
  }
  return images;
}

function formatParagraphItems(paragraphs: ParagraphItem[]): ParagraphItem[] {
  return paragraphs.map((p) => ({
    text: p.text.map((t, index) => ({
      text: index === 0 ? t.text[0].toUpperCase() + t.text.slice(1) : t.text,
      isBold: t.isBold,
    })),
    isList: p.isList,
  }));
}

let slideGrids: GridStyle[] = [];
async function getSlideGrids() {
  try {
    slideGrids = await figma.getLocalGridStylesAsync();
  } catch (error) {
    console.error('Error fetching slide grids:', error);
  }
}

async function renderSlidesToFrames(
  slides: SlideData[],
  zip: JSZip,
  name: string,
  posX: number
) {
  const section = createSection(name, posX);
  const parentFrame = createParentFrame(section);
  parentFrame.fills = [];

  for (const [index, slide] of slides.entries()) {
    slide.number = index + 1;
    const frame = await createSlideFrame(parentFrame, slide.number);
    await createTitleFrame(frame, slide.title, slide.number);
    const bodyFrame = createBodyFrame(frame, slide.number);
    await createTextFrames(bodyFrame, slide.text, slide.number);
    await createImageFrames(bodyFrame, slide.images, zip);
  }

  section.resizeWithoutConstraints(2120, slides.length * 1140 + 140);
  parentFrame.relativeTransform = [
    [1, 0, 100],
    [0, 1, 100],
  ];
}

function createSection(name: string, posX: number) {
  const section = figma.createSection();
  figma.currentPage.appendChild(section);
  section.name = name?.match(/[0-9]/g)?.join('') ?? 'Slides';
  section.resizeWithoutConstraints(2120, 1080);
  section.x = posX;
  return section;
}

function createParentFrame(section: SectionNode) {
  const parentFrame = figma.createFrame();
  section.appendChild(parentFrame);
  parentFrame.layoutMode = 'VERTICAL';
  parentFrame.layoutSizingHorizontal = 'HUG';
  parentFrame.layoutSizingVertical = 'HUG';
  parentFrame.itemSpacing = 60;
  return parentFrame;
}

async function createSlideFrame(parentFrame: FrameNode, number: number) {
  const frame = figma.createFrame();
  parentFrame.appendChild(frame);
  frame.name = `${number}`;
  await frame.setGridStyleIdAsync(slideGrids[0].id);
  frame.resize(1920, 1080);
  frame.layoutMode = 'VERTICAL';
  frame.layoutSizingHorizontal = 'FIXED';
  frame.layoutSizingVertical = 'FIXED';
  frame.verticalPadding = 80;
  frame.horizontalPadding = 80;
  frame.itemSpacing = 30;

  if (number === 1) {
    frame.primaryAxisAlignItems = 'CENTER';
  }

  return frame;
}

async function createTitleFrame(
  frame: FrameNode,
  title: string,
  number: number
) {
  const titleFrame = figma.createFrame();
  frame.appendChild(titleFrame);
  titleFrame.layoutMode = 'VERTICAL';
  titleFrame.layoutSizingHorizontal = 'FILL';
  titleFrame.layoutSizingVertical = 'HUG';
  titleFrame.itemSpacing = 25;

  const titleText = figma.createText();
  titleFrame.appendChild(titleText);
  await figma.loadFontAsync({ family: 'Roboto', style: 'ExtraBold' });
  await figma.loadFontAsync({ family: 'Roboto', style: 'Bold' });
  if (number === 1) {
    titleText.fontName = { family: 'Roboto', style: 'Bold' };
  } else {
    titleText.fontName = { family: 'Roboto', style: 'ExtraBold' };
  }
  titleText.characters = title;
  if (number === 1) {
    titleText.fontSize = 64;
    titleText.fills = [
      {
        type: 'SOLID',
        color: { r: 0.168, g: 0.212, b: 0.454 },
      },
    ];
  } else {
    titleText.fontSize = 48;
    titleText.fills = [
      {
        type: 'SOLID',
        color: { r: 0.176, g: 0.212, b: 0.455 },
      },
    ];
  }

  titleText.lineHeight = {
    value: 140,
    unit: 'PERCENT',
  };
  titleText.resizeWithoutConstraints(1920, 1080);
  titleText.layoutSizingHorizontal = 'FILL';
  titleText.layoutSizingVertical = 'HUG';
}

function createBodyFrame(frame: FrameNode, number: number) {
  const bodyFrame = figma.createFrame();
  frame.appendChild(bodyFrame);
  bodyFrame.layoutMode = 'VERTICAL';
  if (number === 1) {
    bodyFrame.layoutSizingVertical = 'HUG';
  } else {
    bodyFrame.layoutSizingVertical = 'FILL';
  }
  bodyFrame.layoutSizingHorizontal = 'FILL';
  bodyFrame.counterAxisAlignItems = 'CENTER';
  bodyFrame.itemSpacing = 25;
  return bodyFrame;
}

async function createTextFrames(
  bodyFrame: FrameNode,
  ParagraphItems: ParagraphItem[],
  number: number
) {
  await figma.loadFontAsync({ family: 'Roboto', style: 'Regular' });
  await figma.loadFontAsync({ family: 'Roboto', style: 'Bold' });

  for (const p of ParagraphItems) {
    const textFrame = figma.createText();
    textFrame.fontName = { family: 'Roboto', style: 'Regular' };
    textFrame.characters = p.text.map((t) => t.text).join('');

    let currentIndex = 0;
    for (const t of p.text) {
      const length = t.text.length;
      if (t.isBold || number === 1) {
        textFrame.setRangeFontName(currentIndex, currentIndex + length, {
          family: 'Roboto',
          style: 'Bold',
        });
      }
      currentIndex += length;
    }

    textFrame.fills = [
      { type: 'SOLID', color: { r: 0.439, g: 0.494, b: 0.682 } },
    ];
    if (number === 1) {
      textFrame.fontSize = 48;
      textFrame.lineHeight = {
        value: 100,
        unit: 'PERCENT',
      };
    } else {
      textFrame.setRangeListOptions(0, textFrame.characters.length, {
        type: p.isList ? 'UNORDERED' : 'NONE',
      });
      textFrame.fontSize = 36;
      textFrame.listSpacing = 25;
      textFrame.lineHeight = {
        value: 120,
        unit: 'PERCENT',
      };
    }

    bodyFrame.appendChild(textFrame);
    textFrame.layoutSizingHorizontal = 'FILL';
    textFrame.layoutSizingVertical = 'HUG';
  }
}

async function createImageFrames(
  bodyFrame: FrameNode,
  images: string[],
  zip: JSZip
) {
  let picFrame;
  if (images.length > 1) {
    picFrame = figma.createFrame();
    picFrame.name = 'picFrame';
    bodyFrame.appendChild(picFrame);
    picFrame.layoutMode = 'HORIZONTAL';
    picFrame.layoutSizingHorizontal = 'HUG';
    picFrame.layoutSizingVertical = 'FILL';
    picFrame.itemSpacing = 30;
  }

  for (const imagePath of images) {
    const imageFile = zip.file(`${imagePath}`);
    if (imageFile) {
      const imageData = await imageFile.async('uint8array');
      const imageType: string = imagePath.split('.').pop() || '';

      if (['png', 'jpeg', 'jpg'].includes(imageType)) {
        const image = figma.createImage(imageData);
        const { width, height } = await image.getSizeAsync();
        const ratio = width / height;

        const imageNode = figma.createRectangle();
        imageNode.cornerRadius = 30;

        picFrame
          ? picFrame.appendChild(imageNode)
          : bodyFrame.appendChild(imageNode);

        imageNode.layoutSizingVertical = 'FILL';
        imageNode.resize(imageNode.height * ratio, imageNode.height);

        imageNode.fills = [
          {
            type: 'IMAGE',
            scaleMode: 'FIT',
            imageHash: image.hash,
          },
        ];
      } else {
        console.error(`Unsupported image type: ${imageType}`);
      }
    } else {
      console.error(`Image file not found for path: ${imagePath}`);
    }
  }
}
