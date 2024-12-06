import JSZip from 'jszip';
import { parseStringPromise } from 'xml2js';

interface TextItem {
  text: string;
  isList: boolean;
  isBold: boolean;
}

interface SlideData {
  title: string;
  text: TextItem[];
  images: string[];
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
    console.log('Zip file successfully loaded.');
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
  console.log(slides);
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
      text: formatTextItems(paragraphs),
      images,
    };
  } catch (error) {
    console.error(`Error processing slide ${path}:`, error);
    return null;
  }
}

function extractTextData(slideData: any): {
  title: string;
  paragraphs: TextItem[];
} {
  let title = '';
  const paragraphs: TextItem[] = [];

  const shapes =
    slideData?.['p:sld']?.['p:cSld']?.[0]?.['p:spTree']?.[0]?.['p:sp'] || [];
  for (const shape of shapes) {
    const textBody = shape?.['p:txBody'];
    if (textBody) {
      const paragraphsInShape = textBody[0]?.['a:p'];
      if (paragraphsInShape) {
        for (const paragraph of paragraphsInShape) {
          const text = (paragraph?.['a:r'] || [])
            .map((r: any) => r['a:t']?.[0] || '')
            .join('')
            .trim();

          const style = paragraph?.['a:r']?.[0]?.['a:rPr']?.[0];
          const isBold = style?.$?.b === '1';

          if (!text) continue;

          const paragraphProps = paragraph?.['a:pPr']?.[0];
          const isList = !paragraphProps || !paragraphProps['a:buNone'];

          if (!title && text) title = text;
          else
            paragraphs.push({
              text,
              isList,
              isBold,
            });
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

function formatTextItems(paragraphs: TextItem[]): TextItem[] {
  return paragraphs.map((p) => ({
    text: `${p.text[0].toUpperCase() + p.text.slice(1)}`,
    isList: p.isList,
    isBold: p.isBold,
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

  for (const [index, slide] of slides.entries()) {
    const frame = await createSlideFrame(parentFrame, index);
    await createTitleFrame(frame, slide.title);
    const bodyFrame = createBodyFrame(frame);
    await createTextFrames(bodyFrame, slide.text);
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

async function createSlideFrame(parentFrame: FrameNode, index: number) {
  const frame = figma.createFrame();
  parentFrame.appendChild(frame);
  frame.name = `${index + 1}`;
  await frame.setGridStyleIdAsync(slideGrids[0].id);
  frame.resize(1920, 1080);
  frame.layoutMode = 'VERTICAL';
  frame.layoutSizingHorizontal = 'FIXED';
  frame.layoutSizingVertical = 'FIXED';
  frame.verticalPadding = 80;
  frame.horizontalPadding = 80;
  frame.itemSpacing = 30;
  return frame;
}

async function createTitleFrame(frame: FrameNode, title: string) {
  const titleFrame = figma.createFrame();
  frame.appendChild(titleFrame);
  titleFrame.layoutMode = 'VERTICAL';
  titleFrame.layoutSizingHorizontal = 'FILL';
  titleFrame.layoutSizingVertical = 'HUG';
  titleFrame.itemSpacing = 25;

  const titleText = figma.createText();
  titleFrame.appendChild(titleText);
  await figma.loadFontAsync({ family: 'Roboto', style: 'ExtraBold' });
  titleText.fontName = { family: 'Roboto', style: 'ExtraBold' };
  titleText.characters = title;
  titleText.fontSize = 48;
  titleText.fills = [
    {
      type: 'SOLID',
      color: { r: 0.176, g: 0.212, b: 0.455 },
    },
  ];
  titleText.lineHeight = {
    value: 140,
    unit: 'PERCENT',
  };
  titleText.resizeWithoutConstraints(1920, 1080);
  titleText.layoutSizingHorizontal = 'FILL';
  titleText.layoutSizingVertical = 'HUG';
}

function createBodyFrame(frame: FrameNode) {
  const bodyFrame = figma.createFrame();
  frame.appendChild(bodyFrame);
  bodyFrame.layoutMode = 'VERTICAL';
  bodyFrame.layoutSizingHorizontal = 'FILL';
  bodyFrame.layoutSizingVertical = 'FILL';
  bodyFrame.itemSpacing = 25;
  return bodyFrame;
}

async function createTextFrames(bodyFrame: FrameNode, textItems: TextItem[]) {
  await figma.loadFontAsync({ family: 'Roboto', style: 'Regular' });
  await figma.loadFontAsync({ family: 'Roboto', style: 'Bold' });

  const mergedTextArray = mergeTextItems(textItems);

  for (const p of mergedTextArray) {
    const textFrame = figma.createText();

    textFrame.fontName = p.isBold
      ? { family: 'Roboto', style: 'Bold' }
      : { family: 'Roboto', style: 'Regular' };
    textFrame.characters = p.text;
    textFrame.setRangeListOptions(0, textFrame.characters.length, {
      type: p.isList ? 'UNORDERED' : 'NONE',
    });
    textFrame.fontSize = 36;
    textFrame.listSpacing = 25;
    textFrame.fills = [
      { type: 'SOLID', color: { r: 0.439, g: 0.494, b: 0.682 } },
    ];
    textFrame.lineHeight = {
      value: 120,
      unit: 'PERCENT',
    };

    bodyFrame.appendChild(textFrame);
    textFrame.layoutSizingHorizontal = 'FILL';
    textFrame.layoutSizingVertical = 'HUG';
  }

  bodyFrame.layoutSizingHorizontal = 'FILL';
  bodyFrame.layoutSizingVertical = 'FILL';
  bodyFrame.counterAxisAlignItems = 'CENTER';
}

function mergeTextItems(textItems: TextItem[]): TextItem[] {
  return textItems.reduce<TextItem[]>((acc, p) => {
    const lastItem = acc[acc.length - 1];
    if (lastItem && lastItem.isList === p.isList) {
      lastItem.text += `\n${p.text}`;
    } else {
      acc.push({ ...p });
    }
    return acc;
  }, []);
}

async function createImageFrames(
  bodyFrame: FrameNode,
  images: string[],
  zip: JSZip
) {
  let picFrame;
  if (images.length > 1) {
    picFrame = figma.createFrame();
    bodyFrame.appendChild(picFrame);
    picFrame.layoutMode = 'HORIZONTAL';
    picFrame.layoutSizingHorizontal = 'FIXED';
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
