import JSZip from "jszip";
import { parseStringPromise } from "xml2js";

interface SlideData {
	title: string;
	text: string[];
	images: string[];
}

interface PluginMessage {
	type: string;
	data: ArrayBuffer;
	name: string;
	posX: number;
}

interface ParagraphData {
	text: string;
	isList: boolean;
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
			const paragraphs: ParagraphData[] = [];
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

							if (!text) continue; // Пропускаем пустые параграфы

							const paragraphProps = paragraph?.["a:pPr"]?.[0];
							const isList =
								!paragraphProps || !paragraphProps["a:buNone"];

							if (!title && text) title = text;
							else paragraphs.push({ text, isList });
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

			// Формируем текст с маркером для списка
			const text: string[] = [];
			for (const p of paragraphs) {
				if (p.isList) {
					text.push(`• ${p.text[0].toUpperCase() + p.text.slice(1)}`);
				} else {
					text.push(p.text);
				}
			}

			slides.push({
				title,
				text,
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
	// section.setFillStyleIdAsync(nullPaint.id);

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

		const bodyText = figma.createText();
		bodyFrame.appendChild(bodyText);
		await figma.loadFontAsync({ family: "Roboto", style: "Regular" });
		bodyText.fontName = { family: "Roboto", style: "Regular" };

		for (const p of slide.text) {
			if (p.startsWith("•") || p.match(/^\d+\. /)) {
				// Если это список, форматируем его как список
				if (p.startsWith("•")) {
					const paragraph = p.substring(2); // Remove the bullet point
					bodyText.characters += `${paragraph}\n`;
					bodyText.setRangeListOptions(
						bodyText.characters.length - paragraph.length - 1,
						bodyText.characters.length - 1,
						{ type: "UNORDERED" }
					);
				} else {
					const match = p.match(/^\d+\. /);
					if (match) {
						const paragraph = p.substring(match[0].length); // Remove the number and period and space
						bodyText.characters += `${paragraph}\n`;
						bodyText.setRangeListOptions(
							bodyText.characters.length - paragraph.length - 1,
							bodyText.characters.length - 1,
							{ type: "ORDERED" }
						);
					}
				}
			} else {
				// Если это не список, просто добавляем его к тексту
				bodyText.characters += `${p}\n`;
				bodyText.setRangeListOptions(
					bodyText.characters.length - p.length - 1,
					bodyText.characters.length - 1,
					{ type: "NONE" }
				);
			}
		}

		bodyText.characters = bodyText.characters.slice(0, -1);
		bodyText.fontSize = 36;
		bodyText.fills = [
			{ type: "SOLID", color: { r: 0.439, g: 0.494, b: 0.682 } },
		];
		bodyText.lineHeight = {
			value: 120,
			unit: "PERCENT",
		};
		bodyText.resize(1920, 1080);
		bodyText.layoutSizingHorizontal = "FILL";
		bodyText.layoutSizingVertical = "HUG";

		let picFrame;
		if (slide.images.length > 1) {
			picFrame = figma.createFrame();
			bodyFrame.appendChild(picFrame);
			picFrame.layoutMode = "HORIZONTAL";
			picFrame.layoutSizingHorizontal = "FILL";
			picFrame.layoutSizingVertical = "FILL";
			picFrame.itemSpacing = 30;
		}
		for (const imagePath of slide.images) {
			const imageFile = zip.file(`${imagePath}`);
			if (imageFile) {
				const imageData = await imageFile.async("uint8array");

				// Логирование типа изображения
				const imageType = imagePath.split(".").pop();
				// console.log(`Loading image: ${imagePath}, Type: ${imageType}`);

				// Проверка поддерживаемых форматов
				if (
					imageType === "png" ||
					imageType === "jpeg" ||
					imageType === "jpg"
				) {
					const image = figma.createImage(imageData);
					const imageNode = figma.createRectangle();
					picFrame
						? picFrame.appendChild(imageNode)
						: bodyFrame.appendChild(imageNode);
					imageNode.cornerRadius = 30;
					imageNode.layoutSizingHorizontal = "FILL";
					imageNode.layoutSizingVertical = "FILL";

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
