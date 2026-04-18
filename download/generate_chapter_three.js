const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Header, Footer,
  AlignmentType, HeadingLevel, PageBreak, PageNumber,
  LevelFormat, BorderStyle, WidthType, Table, TableRow, TableCell,
  VerticalAlign, ShadingType
} = require("docx");

const c = {
  primary: "1B2A4A",
  body: "1E293B",
  secondary: "4A5568",
  accent: "8B4513",
};

function bp(runs, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { after: 160, line: 276 },
    ...opts,
    children: runs.map(r =>
      typeof r === "string"
        ? new TextRun({ text: r, font: "Calibri", size: 22, color: c.body })
        : r
    ),
  });
}

function t(text, extra = {}) {
  return new TextRun({ text, font: "Calibri", size: 22, color: c.body, ...extra });
}
function ti(text, extra = {}) {
  return new TextRun({ text, font: "Calibri", size: 22, color: c.body, italics: true, ...extra });
}
function tb(text, extra = {}) {
  return new TextRun({ text, font: "Calibri", size: 22, color: c.body, bold: true, ...extra });
}

// ===== 3.1 MATERIALS AND METHODS (heading only) =====
// ===== 3.1.1 Materials =====
const s311 = bp([
  t("Petri dishes (90 mm), conical flask (500 ml), beaker (250 ml), measuring cylinder (250 ml), micropipette, foil paper, Bunsen burner, masking tape, spatula, glass slides, coverslips, tweezers, cotton wool, Sabouraud Dextrose Agar (SDA) powder, and chloramphenicol antibiotic powder.")
]);

// ===== 3.1.2 Equipment =====
const s312 = bp([
  t("Hot air oven, analytical weighing balance, autoclave, refrigerator, microscope (compound light microscope with x10 and x40 objectives), incubator, and gas cylinder.")
]);

// ===== 3.1.3 Reagents =====
const s313 = bp([
  t("Lactophenol cotton blue stain.")
]);

// ===== 3.2 Research Design =====
const s32 = bp([
  t("This study adopted a descriptive cross-sectional design aimed at isolating and identifying the fungal microflora present in the indoor air of the reception areas of three academic buildings (Niger, Volta, and Limpopo) at Nile University of Nigeria, Abuja. The research involved the collection of air samples using the passive sedimentation (settle plate) technique, culturing of fungal spores on Sabouraud Dextrose Agar (SDA), and identification of isolates based on their macroscopic colony characteristics. Samples were collected at two different times of the day (morning and afternoon) to capture possible temporal variation in fungal spore loads across the three buildings.")
]);

// ===== 3.3 Study Area =====
const s33 = [
  bp([
    t("The study was conducted at Nile University of Nigeria, located along the Federal Capital Territory Expressway in Abuja, Nigeria. Abuja lies within Nigeria's middle belt, characterised by distinct wet and dry seasons, a mix of Sudan savanna vegetation, and an altitude of approximately 840 metres above sea level. These climatic conditions, combined with warm daytime temperatures and seasonal rainfall, create an environment that can support fungal growth and spore dispersal.")
  ]),
  bp([
    t("Sampling was carried out in the reception areas of three academic buildings on the main campus:")
  ]),
  bp([
    t("1. Niger Building \u2013 houses administrative offices and receives high volumes of foot traffic from students and visitors attending to administrative matters.")
  ], { indent: { left: 360 } }),
  bp([
    t("2. Volta Building \u2013 used heavily for undergraduate lectures, with a reception area near the main entrance.")
  ], { indent: { left: 360 } }),
  bp([
    t("3. Limpopo Building \u2013 serves a mix of postgraduate and general academic functions, with an open reception space near the building entrance.")
  ], { indent: { left: 360 } }),
  bp([
    t("Each reception area was selected because it represents a high-traffic zone where students, staff, and visitors congregate, sometimes for extended periods, while waiting for appointments or attending to administrative tasks.")
  ]),
];

// ===== 3.4 Sample Collection =====
const s34 = [
  bp([
    t("Airborne fungal spores were collected using the sedimentation (settle plate) method (Pasquarella et al., 2000; Viani et al., 2020). Sterile Petri dishes containing Sabouraud Dextrose Agar (SDA) supplemented with chloramphenicol were used to suppress bacterial growth during sampling and incubation (Black, 2020).")
  ]),
  bp([
    t("At each building, four SDA plates were taken to the reception area. Three plates were designated as sample plates and one as a control. The sample plates were placed on a level surface at a height of approximately 1.0 m above the ground, which corresponds roughly to the average window height and is a standard sampling height used in indoor aeromycological surveys (Madukasi et al., 2021). The control plate remained securely closed throughout the exposure period to check for any contamination introduced during media preparation or transport.")
  ]),
  bp([
    t("The lids of the three sample plates were removed simultaneously and exposed to the air for exactly 15 minutes, after which the lids were replaced and each plate was sealed with masking tape. The plates were then placed into sealed bags and transported back to the laboratory in a cool box.")
  ]),
  bp([
    tb("Morning (A.M.) sampling"), t("")
  ], { spacing: { after: 80, line: 276 } }),
  bp([
    t("Morning samples were collected between 9:40 a.m. and 10:15 a.m., beginning with the Niger building, followed by the Volta building, and then the Limpopo building. Ambient temperatures recorded during this period were approximately 34.5 \u00B0C to 34.7 \u00B0C across the three buildings.")
  ]),
  bp([
    tb("Afternoon (P.M.) sampling"), t("")
  ], { spacing: { after: 80, line: 276 } }),
  bp([
    t("The same sampling procedure was repeated in the afternoon. Afternoon temperatures were higher than the morning readings: 36.0 \u00B0C to 36.5 \u00B0C in the Niger building, 36.0 \u00B0C in the Volta building, and 36.3 \u00B0C in the Limpopo building.")
  ]),
];

// ===== 3.5 Preparation of Culture Media =====
const s35 = [
  bp([
    t("Sabouraud Dextrose Agar (SDA) was used as the culture medium for fungal isolation. SDA is a general-purpose medium widely used for the recovery of airborne fungal conidia (Guinea et al., 2005). To prepare the medium, 16.25 g of SDA powder was measured using an analytical weighing balance and dissolved in 250 ml of distilled water in a clean conical flask. The flask was stoppered with cotton wool and covered with aluminium foil.")
  ]),
  bp([
    t("The medium was sterilised by autoclaving at 121 \u00B0C and 15 psi for 15 minutes. After autoclaving, the medium was allowed to cool to approximately 45\u201350 \u00B0C. At this point, chloramphenicol was added aseptically at the manufacturer\u2019s recommended concentration to inhibit bacterial growth (Hocking and Pitt, 1980).")
  ]),
  bp([
    t("The supplemented agar was then poured into sterile 90 mm disposable Petri dishes, with approximately 20\u201325 ml of medium per dish, and allowed to solidify on a level surface. The plates were checked for any signs of contamination and stored in sealed bags at 4 \u00B0C until needed. Fresh plates were prepared on the same day as each sampling visit.")
  ]),
];

// ===== 3.6 Incubation and Isolation of Fungi =====
const s36 = [
  bp([
    t("All plates (sample and control) were incubated in an upright position at 28 \u00B0C for 5 to 7 days. An incubation temperature in the range of 25\u201328 \u00B0C is standard for mesophilic fungi, which make up the majority of indoor airborne species (Barnett and Hunter, 1998; Tomazin and Matos, 2024). The control plates showed no growth in any of the buildings, confirming that contamination during media preparation and transport did not occur.")
  ]),
  bp([
    t("After the incubation period, the visible fungal colonies on each sample plate were counted. The colony counts were recorded separately for each building and each time of day (A.M. and P.M.) as shown below:")
  ]),
  bp([
    t("Volta Building: 11 colonies (A.M.) and 6 colonies (P.M.)")
  ], { indent: { left: 360 } }),
  bp([
    t("Limpopo Building: 10 colonies (A.M.) and 9 colonies (P.M.)")
  ], { indent: { left: 360 } }),
  bp([
    t("Niger Building: 6 colonies (A.M.) and 8 colonies (P.M.)")
  ], { indent: { left: 360 } }),
  bp([
    t("The total colony count across all buildings and both sampling times was 50. For the purpose of obtaining pure cultures, 15 colonies were selected for sub-culturing based on visual differences in colony morphology (colour, texture, and growth pattern) to ensure that the widest possible range of colony types was represented. A fresh batch of SDA was prepared (16.5 g dissolved in the appropriate volume of distilled water, autoclaved, supplemented with chloramphenicol, and poured into 15 Petri dishes).")
  ]),
  bp([
    t("Using tweezers sterilised by flaming in a Bunsen burner and wiped with 70% ethanol, a small portion of each selected colony was transferred onto the surface of a fresh SDA plate. Each sub-culture was placed on its own individual plate, labelled as \u201Cunknown\u201D with a numerical identifier, and incubated at 28 \u00B0C for another 5 to 7 days.")
  ]),
];

// ===== 3.7 Identification and Characterization of Isolates =====
const s37_intro = bp([
  t("Identification of the fungal isolates was based on macroscopic and microscopic characteristics of the pure cultures obtained after sub-culturing.")
]);

// 3.7.1 Macroscopic Examination
const s371 = [
  bp([
    t("After the second incubation period, the sub-cultured plates were examined for colony growth and purity. Of the 15 sub-cultures, 7 produced pure cultures, 3 yielded mixed cultures, and the remaining 5 showed either no growth or colonies too sparse to classify. The pure cultures were examined macroscopically for the following characteristics:")
  ]),
  bp([
    t("Colony colour (obverse and reverse)")
  ], { indent: { left: 360 } }),
  bp([
    t("Texture (velvety, powdery, cottony, or floccose)")
  ], { indent: { left: 360 } }),
  bp([
    t("Colony margin and surface morphology")
  ], { indent: { left: 360 } }),
  bp([
    t("Growth rate")
  ], { indent: { left: 360 } }),
  bp([
    t("Pigment production and soluble pigments")
  ], { indent: { left: 360 } }),
  bp([
    t("Based on these macroscopic features, and with reference to standard identification keys (Klich, 2002; Visagie et al., 2014; Houbraken et al., 2020), the following fungal species were identified among the pure cultures:")
  ]),
  bp([
    t("1. "), ti("Aspergillus niger"), t(" (white strain) \u2013 white, velvety colony with pale reverse.")
  ], { indent: { left: 360 } }),
  bp([
    t("2. "), ti("Aspergillus niger"), t(" (black strain) \u2013 dense, black, granular colony characteristic of the species.")
  ], { indent: { left: 360 } }),
  bp([
    t("3. "), ti("Penicillium simplicissimum"), t(" (black) \u2013 dark greyish-black colony with finely velvety texture and regular margin.")
  ], { indent: { left: 360 } }),
  bp([
    t("4. "), ti("Aspergillus flavus"), t(" \u2013 yellow-green to olive/blue-green colony with pale reverse.")
  ], { indent: { left: 360 } }),
  bp([
    t("5. "), ti("Trichoderma erinaceum"), t(" (white) \u2013 fast-growing, white, cottony colony that spread rapidly across the plate.")
  ], { indent: { left: 360 } }),
  bp([
    t("6. "), ti("Aspergillus melleus"), t(" \u2013 yellowish-brown to tan colony consistent with published descriptions.")
  ], { indent: { left: 360 } }),
  bp([
    t("7. "), ti("Aspergillus japonicus"), t(" (black and white) \u2013 colony showing distinctive two-tone appearance with black and white sectors.")
  ], { indent: { left: 360 } }),
];

// 3.7.2 Microscopic Examination
const s372 = [
  bp([
    t("A small portion of each colony was mounted on a clean glass slide with a drop of lactophenol cotton blue stain and covered with a coverslip. Observations were made under a compound light microscope using the x10 and x40 objectives. Fungal structures such as hyphae (septate or aseptate), conidia, conidiophores, and spore arrangements were observed and compared with standard identification keys (Klich, 2002; Visagie et al., 2014; Houbraken et al., 2020). The microscopic features observed were consistent with the macroscopic identifications. For example, "),
    ti("Aspergillus niger"), t(" showed dark brown to black conidial heads with radiating, compact conidiophore structures, while "),
    ti("Penicillium simplicissimum"), t(" displayed the characteristic brush-like (penicillus) conidiophore arrangement typical of the genus.")
  ]),
  bp([
    t("It should be noted that the identifications presented here are based on morphological features alone. Molecular methods (PCR, DNA sequencing) were not employed in this study. Species-level identification of "),
    ti("Aspergillus"), t(" and "),
    ti("Penicillium"), t(" from colony and microscopic morphology alone carries some degree of uncertainty, and the identifications should be regarded as preliminary.")
  ]),
];

// ===== 3.8 Quality Control =====
const s38 = [
  bp([
    t("To ensure accuracy and reliability of results, the following quality control measures were observed:")
  ]),
  bp([
    t("All glassware and instruments were sterilised before and after use, either by autoclaving or by flaming in a Bunsen burner.")
  ]),
  bp([
    t("Media sterility was confirmed by incubating one unexposed control plate per building alongside the sample plates. No growth was observed on any control plate.")
  ]),
  bp([
    t("Sampling was conducted under aseptic conditions to prevent external contamination of the plates.")
  ]),
  bp([
    t("Duplicate sampling (three plates per location per time of day) was carried out to ensure consistency of results.")
  ]),
];

// ===== REFERENCES =====
const refs = [
  'Barnett, H.L. and Hunter, B.B. (1998). Illustrated Genera of Imperfect Fungi. 4th edition. APS Press, St. Paul, MN.',
  'Black, W.D. (2020). A comparison of several media types and basic techniques used to assess outdoor airborne fungi in Melbourne, Australia. PLoS ONE, 15(12), e0238901.',
  'Guinea, J., Pelaez, T., Alhambra, A. and Bouza, E. (2005). Evaluation of Czapeck agar and Sabouraud dextrose agar for the culture of airborne Aspergillus conidia. Medical Mycology, 43(5), 477\u2013481.',
  'Hocking, A.D. and Pitt, J.I. (1980). Dichloran-glycerol medium for enumeration of xerophilic fungi from low-moisture foods. Applied and Environmental Microbiology, 39(2), 488\u2013492.',
  'Houbraken, J., Frisvad, J.C., Varga, J. and Samson, R.A. (2020). Classification of Aspergillus, Penicillium, Talaromyces and related genera (Eurotiales). Studies in Mycology, 95, 5\u2013169.',
  'Klich, M.A. (2002). Identification of Common Aspergillus Species. Centraalbureau voor Schimmelcultures, Utrecht, The Netherlands.',
  'Madukasi, I., Ezejiofor, A.N., Nwaoguikpe, R.N. and Okolo, B.N. (2021). Microbiological indoor air quality within a tertiary institution in south-east, Nigeria. International Journal of Multi-disciplinary Academic Studies, 9(1), 1\u201314.',
  'Pasquarella, C., Pitzurra, O. and Savino, A. (2000). The Index of Microbial Air Contamination. Journal of Hospital Infection, 46, 241\u2013256.',
  'Tomazin, R. and Matos, T. (2024). Mycological Methods for Routine Air Sampling and Interpretation of Results in Operating Theaters. Diagnostics, 14(3), 288.',
  'Viani, I., Colucci, M.E., Veronesi, L. et al. (2020). Passive air sampling: the use of the index of microbial air contamination. Acta Bio Medica: Atenei Parmensis, 91(3-S), 92\u2013105.',
  'Visagie, C.M., Houbraken, J., Frisvad, J.C. et al. (2014). Identification and nomenclature of the genus Penicillium. Studies in Mycology, 78, 343\u2013371.',
];

// ==================== DOCUMENT ====================
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 22, color: c.body } },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Times New Roman", color: c.primary },
        paragraph: { spacing: { before: 600, after: 300 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Times New Roman", color: c.primary },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullet-list",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }],
      },
    ],
  },
  sections: [
    // ---- COVER ----
    {
      properties: {
        page: {
          margin: { top: 0, bottom: 0, left: 0, right: 0 },
          size: { width: 11906, height: 16838 },
        },
        titlePage: true,
      },
      children: [
        new Paragraph({ spacing: { before: 4800 }, children: [t("")] }),
        new Paragraph({
          alignment: AlignmentType.CENTER, spacing: { after: 400 },
          children: [new TextRun({ text: "\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501", font: "Calibri", size: 20, color: c.accent })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER, spacing: { after: 200 },
          children: [new TextRun({ text: "CHAPTER THREE", font: "Times New Roman", size: 44, bold: true, color: c.primary })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER, spacing: { after: 120 },
          children: [new TextRun({ text: "MATERIALS AND METHODS", font: "Times New Roman", size: 36, color: c.secondary })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER, spacing: { after: 600 },
          children: [new TextRun({ text: "\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501", font: "Calibri", size: 20, color: c.accent })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: "Assessment of Fungal Microflora in the Indoor Air", font: "Times New Roman", size: 28, color: c.primary, italics: true })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: "of Reception Areas in Three Academic Buildings", font: "Times New Roman", size: 28, color: c.primary, italics: true })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER, spacing: { after: 400 },
          children: [new TextRun({ text: "(Niger, Volta, Limpopo) in Nile University of Nigeria", font: "Times New Roman", size: 28, color: c.primary, italics: true })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER, spacing: { after: 120 },
          children: [new TextRun({ text: "Nile University of Nigeria", font: "Times New Roman", size: 26, bold: true, color: c.primary })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER, spacing: { after: 120 },
          children: [new TextRun({ text: "Abuja, Nigeria", font: "Calibri", size: 22, color: c.secondary })],
        }),
      ],
    },
    // ---- CONTENT ----
    {
      properties: {
        page: {
          margin: { top: 1800, bottom: 1440, left: 1440, right: 1440 },
          size: { width: 11906, height: 16838 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT, spacing: { after: 0 },
            children: [new TextRun({ text: "Chapter Three \u2014 Materials and Methods", font: "Calibri", size: 18, color: c.secondary, italics: true })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "\u2014 ", font: "Calibri", size: 18, color: c.secondary }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Calibri", size: 18, color: c.secondary }),
              new TextRun({ text: " \u2014", font: "Calibri", size: 18, color: c.secondary }),
            ],
          })],
        }),
      },
      children: [
        // 3.1
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.1  Materials and Methods" })] }),
        // 3.1.1
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: "3.1.1  Materials" })] }),
        s311,
        // 3.1.2
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: "3.1.2  Equipment" })] }),
        s312,
        // 3.1.3
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: "3.1.3  Reagents" })] }),
        s313,

        // 3.2
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.2  Research Design" })] }),
        s32,

        // 3.3
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.3  Study Area" })] }),
        ...s33,

        // 3.4
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.4  Sample Collection" })] }),
        ...s34,

        // 3.5
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.5  Preparation of Culture Media" })] }),
        ...s35,

        // 3.6
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.6  Incubation and Isolation of Fungi" })] }),
        ...s36,

        // 3.7
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.7  Identification and Characterization of Isolates" })] }),
        s37_intro,
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: "3.7.1  Macroscopic Examination" })] }),
        ...s371,
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: "3.7.2  Microscopic Examination" })] }),
        ...s372,

        // 3.8 Quality Control
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.8  Quality Control" })] }),
        ...s38,

        // References
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "References" })] }),
        ...refs.map(ref =>
          bp([t(ref)], { indent: { left: 720, hanging: 720 }, spacing: { after: 100, line: 276 } })
        ),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/z/my-project/download/Chapter_Three_Materials_and_Methods.docx", buffer);
  console.log("Done.");
});
