import JSZip from "jszip";
import { parseStringPromise } from "xml2js";

interface SlideData {
	title: string;
	text: string;
	images: string[];
}

interface PluginMessage {
	type: string;
	data: ArrayBuffer;
}

figma.showUI(__html__);

figma.ui.onmessage = async (msg: PluginMessage) => {
	if (msg.type === "file") {
		try {
			if (!(msg.data instanceof ArrayBuffer)) {
				console.error("Received data is not an ArrayBuffer.");
				return;
			}

			let zip;
			try {
				zip = await JSZip.loadAsync(msg.data);
				console.log("Zip file successfully loaded.");
			} catch (err) {
				console.error("Error loading zip file:", err);
				return;
			}

			const slides = await parseSlides(zip);
			await renderSlidesToFrames(slides, zip);
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
			console.log(`Parsed slide: ${path}`);

			const relsContent = await relsFile.async("text");
			const relsData = await parseStringPromise(relsContent);
			console.log(`Parsed rels: ${relsPath}`);

			let title = "";
			const paragraphs: string[] = [];
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
							const text =
								paragraph?.["a:r"]?.[0]?.["a:t"]?.[0] || "";
							if (!title && text) title = text;
							else paragraphs.push(text);
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
						console.log(`Image extracted for slide: ${path}`);
					} else {
						console.warn(
							`Image file not found for path: ${imagePath}`
						);
					}
				}
			}

			slides.push({
				title,
				text: paragraphs.join(" "),
				images,
			});
		} catch (slideError) {
			console.error(`Error processing slide ${path}:`, slideError);
		}
	}

	console.log(slides);

	return slides;
}

// const nullPaint = figma.createPaintStyle();

async function renderSlidesToFrames(slides: SlideData[], zip: JSZip) {
	// let yPosition = 0;
	// let xPosition = 0;

	const section = figma.createSection();
	figma.currentPage.appendChild(section);
	section.name = zip.name?.match(/[0-9]/g)?.join("") ?? "Slides";
	console.log(`Created section: ${zip.name}`);
	section.resizeWithoutConstraints(2120, 1080);
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
		titleText.layoutSizingHorizontal = "FILL";
		titleText.layoutSizingHorizontal = "HUG";

		const bodyFrame = figma.createFrame();
		frame.appendChild(bodyFrame);
		bodyFrame.layoutMode = "VERTICAL";
		bodyFrame.layoutSizingHorizontal = "FILL";
		bodyFrame.layoutSizingVertical = "FILL";
		bodyFrame.itemSpacing = 25;

		// Render body text
		const bodyText = figma.createText();
		bodyFrame.appendChild(bodyText);
		await figma.loadFontAsync({ family: "Roboto", style: "Regular" });
		bodyText.fontName = { family: "Roboto", style: "Regular" };
		bodyText.characters = slide.text;
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
					console.error(`Unsupported image type: ${imageType}`);
				}
			} else {
				console.error(`Image file not found for path: ${imagePath}`);
			}
		}
	}

	section.resizeWithoutConstraints(2120, slides.length * 1140 + 60);
}
