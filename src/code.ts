import JSZip from "jszip";
import { parseStringPromise } from "xml2js";

interface textItem {
	text: string;
	isList: boolean;
	textDecoration: string;
	lvl: number;
}
interface SlideData {
	title: string;
	text: textItem[];
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
	if (msg.type === "file") {
		try {
			if (!(msg.data instanceof ArrayBuffer)) {
				console.error("Received data is not an ArrayBuffer.");
				return;
			}

			let zip, name, posX;
			try {
				name = msg.name;
				posX = msg.posX;
				zip = await JSZip.loadAsync(msg.data);
				console.log("Zip file successfully loaded.");
			} catch (err) {
				console.error("Error loading zip file:", err);
				return;
			}

			await getSlideGrids();
			const slides = await parseSlides(zip);
			await renderSlidesToFrames(slides, zip, name, posX);
		} catch (generalError) {
			console.error(
				"An error occurred in the file processing:",
				generalError
			);
		}
	}
};

async function parseSlides(zip: JSZip): Promise<SlideData[]> {
	const slides: SlideData[] = [];
	const slidePaths = Object.keys(zip.files).filter(path =>
		path.startsWith("ppt/slides/slide")
	);
	console.log(`Found ${slidePaths.length} slides.`);

	const relsPaths = Object.keys(zip.files).filter(path =>
		path.startsWith("ppt/slides/_rels/slide")
	);
	console.log(`Found ${relsPaths.length} slide rels.`);

	const slideToRelsMap: { [key: string]: string } = {};

	for (const relsPath of relsPaths) {
		const slideNumber = relsPath.match(/slide(\d+)/)?.[1];
		if (slideNumber) {
			slideToRelsMap[`ppt/slides/slide${slideNumber}.xml`] = relsPath;
		}
	}

	for (const path of slidePaths) {
		const relsPath = slideToRelsMap[path];
		try {
			const slideFile = zip.file(path);
			const relsFile = zip.file(relsPath);

			if (!slideFile) {
				console.warn(`Slide file not found at path: ${path}`);
				continue;
			}

			if (!relsFile) {
				console.warn(`Slide rels file not found at path: ${relsPath}`);
				continue;
			}

			const slideContent = await slideFile.async("text");
			const slideData = await parseStringPromise(slideContent);

			const relsContent = await relsFile.async("text");
			const relsData = await parseStringPromise(relsContent);

			let title = "";
			const paragraphs: textItem[] = [];
			const images: string[] = [];

			const shapes =
				slideData?.["p:sld"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.[
					"p:sp"
				] || [];
			for (const shape of shapes) {
				const textBody = shape?.["p:txBody"];
				if (textBody) {
					const paragraphsInShape = textBody[0]?.["a:p"];
					if (paragraphsInShape) {
						for (const paragraph of paragraphsInShape) {
							const text = (paragraph?.["a:r"] || [])
								.map((r: any) => r["a:t"]?.[0] || "")
								.join("")
								.trim();

							if (!text) continue;

							const paragraphProps = paragraph?.["a:pPr"]?.[0];
							const textDecoration = paragraphProps?.$?.u;

							const lvl = parseInt(
								paragraph?.["a:pPr"]?.[0]?.$.lvl || "0"
							);

							const isList =
								!paragraphProps || !paragraphProps["a:buNone"];
							if (!title && text) title = text;
							else
								paragraphs.push({
									text,
									isList,
									textDecoration,
									lvl,
								});
						}
					}
				}
			}

			const relationships =
				relsData?.["Relationships"]?.["Relationship"] || [];
			for (const relationship of relationships) {
				if (
					relationship.$.Type ===
					"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
				) {
					const imagePath = `ppt/${relationship.$.Target.replace(
						"../",
						""
					)}`;

					if (imagePath) {
						images.push(`${imagePath}`);
					} else {
						console.warn(
							`Image file not found for path: ${imagePath}`
						);
					}
				}
			}

			const textItems: textItem[] = [];
			for (const p of paragraphs) {
				textItems.push({
					text: `${p.text[0].toUpperCase() + p.text.slice(1)}`,
					isList: p.isList,
					textDecoration: p.textDecoration,
					lvl: p.lvl,
				});
			}

			slides.push({
				title,
				text: textItems,
				images,
			});
		} catch (slideError) {
			console.error(`Error processing slide ${path}:`, slideError);
		}
	}

	console.log(slides);

	return slides;
}

let slideGrids: GridStyle[] = [];
async function getSlideGrids() {
	try {
		slideGrids = await figma.getLocalGridStylesAsync();
	} catch (error) {
		console.error("Error fetching slide grids:", error);
	}
}

async function renderSlidesToFrames(
	slides: SlideData[],
	zip: JSZip,
	name: string,
	posX: number
) {
	const section = figma.createSection();
	figma.currentPage.appendChild(section);
	section.name = name?.match(/[0-9]/g)?.join("") ?? "Slides";
	section.resizeWithoutConstraints(2120, 1080);
	section.x = posX;

	const parentFrame = figma.createFrame();
	section.appendChild(parentFrame);
	parentFrame.layoutMode = "VERTICAL";
	parentFrame.layoutSizingHorizontal = "HUG";
	parentFrame.layoutSizingVertical = "HUG";
	parentFrame.itemSpacing = 60;

	for (const [index, slide] of slides.entries()) {
		const frame = figma.createFrame();
		parentFrame.appendChild(frame);
		frame.name = `${index + 1}`;
		frame.setGridStyleIdAsync(slideGrids[0].id);
		frame.resize(1920, 1080);
		frame.layoutMode = "VERTICAL";
		frame.layoutSizingHorizontal = "FIXED";
		frame.layoutSizingVertical = "FIXED";
		frame.verticalPadding = 80;
		frame.horizontalPadding = 80;
		frame.itemSpacing = 30;

		const titleFrame = figma.createFrame();
		frame.appendChild(titleFrame);
		titleFrame.layoutMode = "VERTICAL";
		titleFrame.layoutSizingHorizontal = "FILL";
		titleFrame.layoutSizingVertical = "HUG";
		titleFrame.itemSpacing = 25;

		const titleText = figma.createText();
		titleFrame.appendChild(titleText);
		await figma.loadFontAsync({ family: "Roboto", style: "ExtraBold" });
		titleText.fontName = { family: "Roboto", style: "ExtraBold" };
		titleText.characters = slide.title;
		titleText.fontSize = 48;
		titleText.fills = [
			{
				type: "SOLID",
				color: { r: 0.176, g: 0.212, b: 0.455 },
			},
		];
		titleText.lineHeight = {
			value: 140,
			unit: "PERCENT",
		};
		titleText.resizeWithoutConstraints(1920, 1080);
		titleText.layoutSizingHorizontal = "FILL";
		titleText.layoutSizingVertical = "HUG";

		const bodyFrame = figma.createFrame();
		frame.appendChild(bodyFrame);
		bodyFrame.layoutMode = "VERTICAL";
		bodyFrame.layoutSizingHorizontal = "FILL";
		bodyFrame.layoutSizingVertical = "FILL";
		bodyFrame.itemSpacing = 25;

		await figma.loadFontAsync({ family: "Roboto", style: "Regular" });

		const mergedTextArray = slide.text.reduce<textItem[]>((acc, p) => {
			const lastItem = acc[acc.length - 1];
			if (lastItem && lastItem.isList === p.isList) {
				lastItem.text += `\n${p.text}`;
			} else {
				acc.push({ ...p });
			}
			return acc;
		}, []);

		for (const p of mergedTextArray) {
			const textFrame = figma.createText();
			textFrame.fontName = { family: "Roboto", style: "Regular" };
			textFrame.characters = p.text;
			textFrame.setRangeListOptions(0, textFrame.characters.length, {
				type: p.isList ? "UNORDERED" : "NONE",
			});
			textFrame.fontSize = 36;
			textFrame.listSpacing = 25;
			textFrame.fills = [
				{ type: "SOLID", color: { r: 0.439, g: 0.494, b: 0.682 } },
			];
			textFrame.lineHeight = {
				value: 120,
				unit: "PERCENT",
			};

			// if (p.lvl !== 0) {
			// }

			bodyFrame.appendChild(textFrame);
			textFrame.layoutSizingHorizontal = "FILL";
			textFrame.layoutSizingVertical = "HUG";
		}

		bodyFrame.layoutSizingHorizontal = "FILL";
		bodyFrame.layoutSizingVertical = "FILL";

		let picFrame;
		if (slide.images.length > 1) {
			picFrame = figma.createFrame();
			bodyFrame.appendChild(picFrame);
			picFrame.layoutMode = "HORIZONTAL";
			picFrame.layoutSizingHorizontal = "FIXED";
			picFrame.layoutSizingVertical = "FILL";
			picFrame.itemSpacing = 30;
		}
		for (const imagePath of slide.images) {
			const imageFile = zip.file(`${imagePath}`);
			if (imageFile) {
				const imageData = await imageFile.async("uint8array");

				const imageType = imagePath.split(".").pop();

				if (
					imageType === "png" ||
					imageType === "jpeg" ||
					imageType === "jpg"
				) {
					const image = figma.createImage(imageData);
					const { width, height } = await image.getSizeAsync();
					const ratio = width / height;

					const imageNode = figma.createRectangle();
					imageNode.cornerRadius = 30;

					picFrame
						? picFrame.appendChild(imageNode)
						: bodyFrame.appendChild(imageNode);

					imageNode.layoutSizingVertical = "FILL";
					imageNode.resize(
						imageNode.height * ratio,
						imageNode.height
					);

					imageNode.fills = [
						{
							type: "IMAGE",
							scaleMode: "FIT",
							imageHash: image.hash,
						},
					];
				} else {
					// console.error(`Unsupported image type: ${imageType}`);
				}
			} else {
				console.error(`Image file not found for path: ${imagePath}`);
			}
		}
	}

	section.resizeWithoutConstraints(2120, slides.length * 1140 + 140);
	parentFrame.relativeTransform = [
		[1, 0, 100],
		[0, 1, 100],
	];
}
