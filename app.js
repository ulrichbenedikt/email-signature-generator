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
const PORT = 3000;

// Middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static("public"));
app.use("/images", express.static(path.join(__dirname, "images")));

//styles
function linkStyle() {
	return {
		color: "0000FF",
		underline: { type: "single" },
	};
}

function boldStyle() {
	return {
		color: "0000FF",
		bold: true,
	};
}
// Route
app.post("/generate", async (req, res) => {
	const {
		firstname,
		lastname,
		job,
		email,
		phone,
		telephone,
		linkedin,
		image,
		booking,
	} = req.body;

	const paragraphs = [
		new Paragraph({
			children: [
				new TextRun({
					text: `${firstname} ${lastname}`,
					...boldStyle(),
				}),
			],
		}),
		new Paragraph(job),
		new Paragraph(""),
		new Paragraph({
			children: [
				new TextRun("email = "),
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
				new TextRun("phone = "),
				new ExternalHyperlink({
					children: [
						new TextRun({
							text: telephone,
							//hyperlink: `mailto:${email}`,
							...linkStyle(),
						}),
					],
					link: `tel:${telephone}`,
				}),
			],
		}),
	];

	if (phone && phone.trim() !== "") {
		paragraphs.push(
			new Paragraph({
				children: [
					new TextRun("mobile = "),
					new ExternalHyperlink({
						children: [
							new TextRun({
								text: phone,
								//hyperlink: `mailto:${email}`,
								...linkStyle(),
							}),
						],
						link: `tel:${phone}`,
					}),
				],
			})
		);
	}
	if (linkedin && linkedin.trim() !== "") {
		paragraphs.push(
			new Paragraph({
				children: [
					new TextRun("linkedIn = "),
					new ExternalHyperlink({
						children: [
							new TextRun({
								text: linkedin.replace(
									"https://linkedin.com/in",
									""
								),
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
					new ExternalHyperlink({
						children: [
							new TextRun({
								text: "//> einen Termin vereinbaren",
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

	if (!fs.existsSync(imagePath)) {
		return res.status(400).send("Image not found or not selected.");
	}

	const imageBuffer = fs.readFileSync(imagePath);
	paragraphs.push(
		new Paragraph({
			children: [
				new ImageRun({
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
					...boldStyle(),
				}),
				new TextRun(" | EDV Dienstleistung und Softwareentwicklung"),
			],
		}),
		new Paragraph("Im Pinderpark 5, 90513 Zirndorf"),
		new Paragraph({
			children: [
				new TextRun("website = "),
				new ExternalHyperlink({
					children: [
						new TextRun({
							text: "puresolution.de",
							...linkStyle(),
						}),
					],
					link: "https://www.puresolution.de",
				}),
			],
		}),
		new Paragraph({
			children: [
				new TextRun("phone = "),
				new ExternalHyperlink({
					children: [
						new TextRun({
							text: "+49 911 81000-0",
							...linkStyle(),
						}),
					],
					link: "tel:+49911810000",
				}),
			],
		}),
		new Paragraph({
			children: [
				new TextRun("fax = "),
				new ExternalHyperlink({
					children: [
						new TextRun({
							text: "+49 911 81000-20",
							...linkStyle(),
						}),
					],
					link: "fax:+499118100020",
				}),
			],
		}),
		new Paragraph(""),
		new Paragraph({
			children: [
				new TextRun({
					text: "PureSolution GmbH | Geschäftsführer: Hasso Leeder, Tilo Siegler | Registergericht Fürth HRB 8044 | USt.-IdNr.: DE209913453 | Datenschutzhinweis: PureSolution GmbH verarbeitet Ihre Kontaktdaten elektronisch. | Weitere Informationen finden Sie in der ",
					size: 16, // 9pt
					color: "BEC0BF",
				}),
				new ExternalHyperlink({
					children: [
						new TextRun({
							text: "Datenschutzerklärung.",
							underline: { type: "single" },
                            size: 16, // 9pt
					        color: "BEC0BF",
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
