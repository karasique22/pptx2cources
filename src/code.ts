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
			await renderSlidesToFrames(slides);
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

	for (const path of slidePaths) {
		try {
			const slideFile = zip.file(path);
			if (!slideFile) {
				console.warn(`Slide file not found at path: ${path}`);
				continue;
			}

			const slideContent = await slideFile.async("text");
			const slideData = await parseStringPromise(slideContent);
			console.log(`Parsed slide: ${path}`);

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
							paragraphs.push(text);
						}
					}
				}
			}

			const pictures =
				slideData?.["p:sld"]?.["p:cSld"]?.[0]?.["p:spTree"]?.[0]?.[
					"p:pic"
				] || [];
			for (const picture of pictures) {
				const blip = picture?.["a:blipFill"]?.[0]?.["a:blip"];
				const imageId = blip?.[0]?.["$"]?.["r:embed"];
				const imagePath = `ppt/media/${imageId}.jpeg`;
				const imageFile = imageId ? zip.file(imagePath) : null;

				if (imageFile) {
					const imageData = await imageFile.async("base64");
					images.push(`data:image/jpeg;base64,${imageData}`);
					console.log(`Image extracted for slide: ${path}`);
				} else {
					console.warn(
						`Image file not found or imageId is missing for path: ${imagePath}`
					);
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

async function renderSlidesToFrames(slides: SlideData[]) {
	let xPosition = 0;
	let yPosition = 0;

	for (const [index, slide] of slides.entries()) {
		const frame = figma.createFrame();
		frame.resize(1920, 1080);
		frame.x = xPosition;
		frame.y = yPosition;
		frame.fills = [{ type: "SOLID", color: { r: 1, g: 1, b: 1 } }];
		figma.currentPage.appendChild(frame);

		// Load and apply title font
		(async () => {
			const titleText = figma.createText();
			await figma.loadFontAsync({ family: "Roboto", style: "ExtraBold" });
			titleText.fontName = { family: "Roboto", style: "ExtraBold" };
			titleText.characters = slide.title;
			titleText.fontSize = 48;
			titleText.fills = [
				{ type: "SOLID", color: { r: 43, g: 54, b: 116 } },
			];
			titleText.x = 80; // Margin from left
			titleText.y = 80; // Margin from top
			frame.appendChild(titleText);
		})();

		async () => {
			await figma.loadFontAsync({ family: "Roboto", style: "Regular" });
			const bodyText = figma.createText();
			bodyText.fontName = { family: "Roboto", style: "Regular" };
			bodyText.characters = slide.text;
			bodyText.fontSize = 36;
			bodyText.fills = [{ type: "SOLID", color: { r: 0, g: 0, b: 0 } }];
			bodyText.x = 80;
			bodyText.y = bodyText.y + bodyText.height + 40;
			frame.appendChild(bodyText);
		};

		let imageYPosition = frame.y + frame.height + 40;
		for (const imageData of slide.images) {
			const image = figma.createImage(
				Uint8Array.from(atob(imageData.split(",")[1]), c =>
					c.charCodeAt(0)
				)
			);
			const imageNode = figma.createRectangle();
			imageNode.resize(400, 300);
			imageNode.fills = [
				{ type: "IMAGE", scaleMode: "FILL", imageHash: image.hash },
			];
			imageNode.x = 80;
			imageNode.y = imageYPosition;
			frame.appendChild(imageNode);

			imageYPosition += imageNode.height + 20;
		}

		yPosition += 1080 + 60;

		if (index === slides.length - 1) {
			yPosition = 0;
			xPosition += 1920 + 240;
		}
	}
}
