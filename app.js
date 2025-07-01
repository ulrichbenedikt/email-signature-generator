// app.js
const express = require("express");
const path = require("path");
const bodyParser = require("body-parser");
const fs = require("fs");
const {
	Document,
	Packer,
	Paragraph,
	TextRun,
	ImageRun,
	ExternalHyperlink,
} = require("docx");

const app = express();
const PORT = 3001;

// Middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static("public"));
app.use("/images", express.static(path.join(__dirname, "images")));

const icons = {
	linkedin: fs.readFileSync(path.join(__dirname, "icons/linkedin.png")),
	email: fs.readFileSync(path.join(__dirname, "icons/email.png")),
	phone: fs.readFileSync(path.join(__dirname, "icons/telefon.png")),
	mobil: fs.readFileSync(path.join(__dirname, "icons/mobiltelefon.png")),
	schedule: fs.readFileSync(path.join(__dirname, "icons/zeitplan.png")),
};

//styles
function linkStyle() {
	return {
		color: "0000FF",
		underline: { type: "single" },
	};
}
function Image(icon){
	return new ImageRun({
		type: "png",
		data: icon,
		transformation: {
			width: 10,
			height: 10,
		}
	})
}

// Route
app.post("/generate", async (req, res) => {
	const {
		firstname,
		lastname,
		job,
		email,
		phone,
		ending,
		linkedin,
		image,
		booking,
	} = req.body;

	const paragraphs = [
		new Paragraph(`${firstname} ${lastname}`),
		new Paragraph(job),
		new Paragraph(
			"____________________________________________________________________"
		),
		new Paragraph(""),
		new Paragraph({
			children: [
				Image(icons.email),
				new TextRun("  "),
				new ExternalHyperlink({
					children: [
						new TextRun({
							text: email,
							//hyperlink: `mailto:${email}`,
							...linkStyle(),
						}),
					],
					link: `mailto:${email}`,
				}),
			],
		}),
		new Paragraph({
			children: [
				Image(icons.phone),
				new TextRun("  "),
				new ExternalHyperlink({
					children: [
						new TextRun({
							text: `+49 911 81000-${ending}`,
							...linkStyle(),
						}),
					],
					link: `tel:+4991181000${ending}`,
				}),
			],
		}),
	];

	if (phone && phone.trim() !== "") {
		paragraphs.push(
			new Paragraph({
				children: [
					Image(icons.mobil),
					new TextRun("  "),
					new ExternalHyperlink({
						children: [
							new TextRun({
								text: phone,
								//hyperlink: `mailto:${email}`,
								...linkStyle(),
							}),
						],
						link: `tel:${phone.replaceAll(" ", "")}`,
					}),
				]
			})
		);
	}
	if (linkedin && linkedin.trim() !== "") {
		paragraphs.push(
			new Paragraph({
				children: [
					Image(icons.linkedin),
					new TextRun("  "),
					new ExternalHyperlink({
						children: [
							new TextRun({
								text: linkedin.replace(/[^ ]*\/in\b/, "/in"),
								//hyperlink: `mailto:${email}`,
								...linkStyle(),
							}),
						],
						link: linkedin,
					}),
				],
			})
		);
	}
	if (booking && booking.trim() !== "") {
		paragraphs.push(
			new Paragraph({
				children: [
					Image(icons.schedule),
					new TextRun("  "),
					new ExternalHyperlink({
						children: [
							new TextRun({
								text: "hier vereinbaren",
								bold: true,
								color: "0000FF",
								underline: { type: "single" },
							}),
						],
						link: booking,
					}),
				],
			})
		);
	}

	paragraphs.push(new Paragraph(""));

	// Add image
	const imagePath = path.join(__dirname, "images", image);

	if (!/^[\w\-]+\.(png|jpg|jpeg)$/i.test(image)) {
		return res.status(400).send("Invalid image filename.");
	}

	if (!fs.existsSync(imagePath)) {
		return res.status(400).send("Image not found or not selected.");
	}

	const imageBuffer = fs.readFileSync(imagePath);
	paragraphs.push(
		new Paragraph({
			children: [
				new ImageRun({
					type: "png",
					data: imageBuffer,
					transformation: { width: 500, height: 155 },
				}),
			],
		})
	);

	paragraphs.push(
		new Paragraph(""),
		new Paragraph({
			children: [
				new TextRun({
					text: "PureSolution GmbH",
					bold: true,
				}),
				new TextRun(" | EDV Dienstleistung und Softwareentwicklung"),
			],
		}),
		new Paragraph({
			children: [
				new TextRun("Im Pinderpark 5, 90513 Zirndorf | "),
				new ExternalHyperlink({
					children: [
						new TextRun({
							text: "www.puresolution.de",
							...linkStyle(),
						}),
					],
					link: "https://www.puresolution.de",
				}),
			],
		}),
		new Paragraph({
			children: [
				new TextRun("tel: "),
				new ExternalHyperlink({
					children: [
						new TextRun({
							text: "+49 911 81000-0",
							...linkStyle(),
						}),
					],
					link: "tel:+49911810000",
				}),
				new TextRun(" | fax: "),
				new ExternalHyperlink({
					children: [
						new TextRun({
							text: "+49 911 81000-20",
							...linkStyle(),
						}),
					],
					link: "tel:+499118100020",
				}),
			],
		}),
		new Paragraph(""),
		new Paragraph({
			children: [
				new TextRun({
					text: "PureSolution GmbH | Geschäftsführer: Hasso Leeder, Tilo Siegler | Registergericht Fürth HRB 8044 |",
					size: 16, // 9pt
					color: "808080",
				}),
			],
		}),
		new Paragraph({
			children: [
				new TextRun({
					text: "USt.-IdNr.: DE209913453 | Datenschutzhinweis: PureSolution GmbH verarbeitet Ihre Kontaktdaten elektronisch. |",
					size: 16, // 9pt
					color: "#808080",
				}),
			],
		}),
		new Paragraph({
			children: [
				new TextRun({
					text: "Weitere Informationen finden Sie in der ",
					size: 16, // 9pt
					color: "#808080",
				}),
				new ExternalHyperlink({
					children: [
						new TextRun({
							text: "Datenschutzerklärung.",
							underline: { type: "single" },
							size: 16, // 9pt
							color: "#808080",
						}),
					],
					link: "https://www.puresolution.de/datenschutz/Datenschutzhinweis_Kontakt.pdf",
				}),
			],
		})
	);

	const doc = new Document({
		creator: `${firstname} ${lastname}`,
		title: `Email-Signatur-${firstname}-${lastname}`,
		description: `Email signature generated document for ${firstname} ${lastname}`,
		sections: [
			{
				children: paragraphs,
			},
		],
		styles: {
			document: {
				run: {
					font: "Calibri",
					size: 22, // 11pt
					color: "003C3C",
				},
				paragraph: {
					spacing: { line: 276 }, // single line spacing
				},
			},
			paragraphStyles: [
				{
					id: "Normal",
					name: "Normal",
					quickFormat: true,
					run: {
						font: "Calibri",
						color: "003C3C",
						size: 22,
					},
				},
				{
					id: "Bold",
					name: "Bold",
					basedOn: "Normal",
					next: "Normal",
					quickFormat: true,
					run: {
						color: "0000FF",
						bold: true,
					},
				},
				{
					id: "Hyperlink",
					name: "Hyperlink",
					basedOn: "Normal",
					next: "Normal",
					quickFormat: true,
					run: {
						color: "0000FF",
						underline: { type: "single" },
					},
				},
			],
		},
	});

	const buffer = await Packer.toBuffer(doc);
	res.setHeader(
		"Content-Disposition",
		"attachment; filename=email-signatur-" +
			firstname +
			"-" +
			lastname +
			".docx"
	);
	res.setHeader(
		"Content-Type",
		"application/vnd.openxmlformats-officedocument.wordprocessingml.document"
	);
	res.send(buffer);
});

// Start server
app.listen(PORT, () => {
	console.log(`Server running at http://localhost:${PORT}`);
});
