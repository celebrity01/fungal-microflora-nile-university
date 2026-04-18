const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Header, Footer,
  AlignmentType, HeadingLevel, PageBreak, PageNumber, BorderStyle
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
    spacing: { after: 160, line: 250 },
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

// ==================== SECTION 3.1 ====================
const s31 = [
  bp([
    t("The study was carried out at Nile University of Nigeria, located along the Federal Capital Territory Expressway in Abuja. Three academic buildings were selected for sampling: the Niger building, the Volta building, and the Limpopo building. These buildings are part of the university's main academic zone and house lecture rooms, offices, and reception areas that receive daily foot traffic from students, academic staff, and administrative personnel. Each building has at least one main reception area near its entrance where visitors and students wait or queue for services. It is these reception areas that served as the sampling points for this study.")
  ]),

  bp([
    t("The three buildings differ somewhat in their layout and usage patterns, which was part of the reason for including all three rather than selecting just one. The Niger building handles a relatively high volume of administrative traffic, the Volta building is used heavily for undergraduate lectures, and the Limpopo building serves a mix of postgraduate and general academic functions. Details of the specific sampling locations within each building are provided in the subsections that follow.")
  ]),
];

// ==================== SECTION 3.2 ====================
const s32 = [
  bp([
    t("The culture medium used throughout this study was Sabouraud Dextrose Agar (SDA). SDA is a general-purpose medium widely used for the isolation and cultivation of fungi, and it has been shown to perform well for recovering airborne fungal conidia in indoor and outdoor settings (Guinea et al., 2005; Black, 2020). It contains dextrose as the carbohydrate source and peptone as the nitrogen source, with a slightly acidic pH (around 5.6) that discourages most bacterial growth while supporting a broad range of fungal species (Pitt and Hocking, 2009).")
  ]),

  bp([
    t("To prepare the medium, 16.25 g of SDA powder was weighed using an analytical balance and dissolved in 250 ml of distilled water in a clean conical flask. The flask was stoppered with a cotton wool plug and aluminium foil, then sterilised in an autoclave at 121 \u00B0C and 15 psi for 15 minutes. After autoclaving, the medium was removed and allowed to cool to roughly 45\u201350 \u00B0C \u2014 warm enough to remain liquid but cool enough to handle without heat damage to the Petri dishes. At this point, chloramphenicol was added at the manufacturer\u2019s recommended concentration to further suppress bacterial growth. The use of chloramphenicol-supplemented SDA for airborne fungal sampling is common practice, as it allows fungal colonies to develop without being overgrown by bacteria (Hocking and Pitt, 1980; Ghazanfari et al., 2023).")
  ]),

  bp([
    t("The supplemented agar was then poured into sterile 90 mm disposable Petri dishes, with approximately 20\u201325 ml of medium per dish, and left undisturbed on a level surface to solidify. Once set, the plates were checked for contamination (turbidity, bubbles, or unintended colony growth) before being stored in sealed bags at 4 \u00B0C until needed. Fresh plates were prepared on the same day as each sampling visit to minimise the risk of contamination during storage.")
  ]),
];

// ==================== SECTION 3.3 ====================
const s33 = [
  bp([
    t("Airborne fungi were collected using the passive sedimentation method, also known as the settle plate or gravity sampling technique. In this method, open Petri dishes are left exposed to the air for a defined period, and airborne particles \u2014 including fungal spores \u2014 settle onto the agar surface by gravity. The method has been in use for decades and remains one of the most common approaches for assessing indoor fungal contamination, particularly in settings where active air samplers are not available (Pasquarella et al., 2000; Viani et al., 2020; Awad and Mawla, 2012). While it does not provide a direct volumetric measurement of airborne spore concentration, it gives a reliable indication of the culturable fungi present in the air and is well suited to comparative studies across different indoor locations (Zhang et al., 2022).")
  ]),

  // Plate placement
  bp([
    t("At each building, four SDA plates were taken to the reception area. Three of these were designated as sample plates and one as a control. The sample plates were placed on a level surface at a height of approximately 1.0 m above the ground \u2014 roughly window height, which is the standard sampling height used in many indoor aeromycological surveys (Madukasi et al., 2021). The control plate was kept securely closed throughout the exposure period to check for any contamination introduced during media preparation, transport, or handling. The lids of the three sample plates were then removed simultaneously, and a timer was started.")
  ]),

  // Exposure time
  bp([
    t("The exposure time was 15 minutes for all plates at all locations. After 15 minutes, the lids were replaced, and each plate was sealed with adhesive tape to prevent accidental opening during transport. The plates were then placed back into their sealed storage bags, kept upright, and transported to the laboratory in a cool box.")
  ]),

  // AM sampling
  bp([
    tb("Morning sampling (A.M.)"), t("")
  ], { spacing: { after: 80, line: 250 } }),
  bp([
    t("Morning samples were collected between 9:40 a.m. and 10:15 a.m. across the three buildings. The order of sampling was Niger, then Volta, then Limpopo. Ambient temperatures during the morning collection period were recorded at each site using a digital thermometer. Morning temperatures were fairly consistent across the three buildings, ranging from approximately 34.5 \u00B0C to 34.7 \u00B0C. Relative humidity was not measured at the time of sampling, which is acknowledged as a limitation of this study.")
  ]),

  // PM sampling
  bp([
    tb("Afternoon sampling (P.M.)"), t("")
  ], { spacing: { after: 80, line: 250 } }),
  bp([
    t("The same procedure was repeated in the afternoon, again starting with the Niger building and moving through to Limpopo. Afternoon temperatures were noticeably higher than the morning readings: 36.0 \u00B0C to 36.5 \u00B0C in the Niger building, 36.0 \u00B0C in the Volta building, and 36.3 \u00B0C in the Limpopo building. The temperature difference between the A.M. and P.M. collections is worth noting, as temperature is a known factor influencing airborne fungal spore levels, though the extent of its effect in this particular study is difficult to determine from a single day\u2019s sampling (Kallawicha et al., 2017).")
  ]),

  // Incubation
  bp([
    tb("Incubation"), t("")
  ], { spacing: { after: 80, line: 250 } }),
  bp([
    t("All plates \u2014 both sample and control \u2014 were incubated in an upright position at 28 \u00B0C for 5 to 7 days. An incubation temperature in the range of 25\u201328 \u00B0C is standard for mesophilic fungi, which make up the vast majority of indoor airborne species (Barnett and Hunter, 1998; Tomazin and Matos, 2024). The control plates showed no growth in any of the buildings, confirming that contamination during media preparation and transport was not a factor.")
  ]),
];

// ==================== SECTION 3.4 ====================
const s34 = [
  bp([
    t("After the 5 to 7 day incubation period, the plates were removed from the incubator and the visible fungal colonies on each plate were counted. Colony counts were recorded separately for each building, each time of day (A.M. and P.M.), and each of the three replicate plates. The counts are summarised below.")
  ], { spacing: { after: 120, line: 250 } }),

  bp([
    tb("Colony counts per building and time of day"), t("")
  ], { spacing: { after: 80, line: 250 } }),
  bp([t("\u2022  Volta building: 11 colonies (A.M.), 6 colonies (P.M.)")], { indent: { left: 360 } }),
  bp([t("\u2022  Limpopo building: 10 colonies (A.M.), 9 colonies (P.M.)")], { indent: { left: 360 } }),
  bp([t("\u2022  Niger building: 6 colonies (A.M.), 8 colonies (P.M.)")], { indent: { left: 360 } }),

  bp([
    t("The total colony count across all buildings and both sampling times was 50. The Volta building had the highest A.M. count (11 colonies) but one of the lower P.M. counts (6). The Limpopo building showed relatively consistent counts across both times of day (10 and 9). The Niger building had the fewest colonies in the morning (6) but more in the afternoon (8). These differences between A.M. and P.M. counts are interesting, though a single round of sampling does not allow for any firm conclusions about temporal patterns.")
  ]),

  // Sub-culturing
  bp([
    tb("Sub-culturing"), t("")
  ], { spacing: { after: 80, line: 250 } }),
  bp([
    t("In order to obtain pure cultures for identification, a sub-culturing step was carried out. A fresh batch of SDA was prepared (16.5 g dissolved in the appropriate volume of distilled water, autoclaved, supplemented with chloramphenicol, and poured into 15 Petri dishes) following the same procedure described in Section 3.2.")
  ]),
  bp([
    t("From the original sample plates, 15 colonies were selected for sub-culturing. The selection was based on visible differences in colony morphology \u2014 colour, texture, growth rate, and margin shape \u2014 so that the widest possible range of colony types was represented. Using a pair of metal tweezers that had been sterilised by flaming in a Bunsen burner and wiped with 70% ethanol, a small portion of each selected colony was picked up and transferred onto the surface of a fresh SDA plate. Each sub-culture was made onto its own individual plate. The new plates were labelled \u201Cunknown\u201D with a numerical identifier and incubated at 28 \u00B0C for another 5 to 7 days.")
  ]),
];

// ==================== SECTION 3.5 ====================
const s35 = [
  bp([
    t("After the second incubation period, the sub-cultured plates were examined for colony growth and purity. Of the 15 sub-cultures, 7 produced what appeared to be pure cultures (a single, uniform colony type covering the plate), while 3 yielded mixed cultures (more than one distinct colony type on the same plate). The remaining 5 plates either showed no growth or produced colonies that were too sparse or ambiguous to classify, and were excluded from further identification. Macroscopic characteristics \u2014 colony diameter, texture (velvety, powdery, or floccose), surface colour, reverse colour, and the presence or absence of soluble pigments \u2014 were recorded for each plate.")
  ]),

  bp([
    t("Identification was carried out to species level where possible, using standard taxonomic keys and reference works. For "),
    ti("Aspergillus"), t(" species, the primary reference was Klich (2002), supplemented by the taxonomic framework in Houbraken et al. (2020). For "),
    ti("Penicillium"), t(" identification, Visagie et al. (2014) was the main reference. The colony descriptions below are based on the macroscopic features observed on SDA after 5 days of incubation at 28 \u00B0C.")
  ]),

  bp([
    tb("Identified species"), t("")
  ], { spacing: { after: 80, line: 250 } }),

  bp([
    t("1.  "),
    ti("Aspergillus niger"), t(" (white strain) \u2014 White, velvety colony with relatively slow expansion on SDA. Reverse pale. Conidial heads not yet fully developed at the time of examination. Distinguished from other white "),
    ti("Aspergillus"), t(" species by colony texture and the pattern of conidiation as described by Klich (2002).")
  ], { indent: { left: 360 }, spacing: { after: 120, line: 250 } }),

  bp([
    t("2.  "),
    ti("Aspergillus niger"), t(" (black strain) \u2014 Dense, black, granular colony characteristic of the species. Reverse also darkened. This is one of the most commonly encountered airborne fungi in indoor environments and was identifiable by its distinctive black conidial heads on SDA (Klich, 2002; Houbraken et al., 2020).")
  ], { indent: { left: 360 }, spacing: { after: 120, line: 250 } }),

  bp([
    t("3.  "),
    ti("Penicillium simplicissimum"), t(" (black) \u2014 Colony with a dark, greyish-black appearance, unusual for the genus. The identification was based on colony texture (finely velvety), margin (regular), and the absence of the typical blue-green conidial colour that characterises most "),
    ti("Penicillium"), t(" species. Visagie et al. (2014) note that some "),
    ti("Penicillium"), t(" species in Section "),
    ti("Robsamsonia"), t(" can produce darker colonies, and "),
    ti("P. simplicissimum"), t(" is one such species.")
  ], { indent: { left: 360 }, spacing: { after: 120, line: 250 } }),

  bp([
    t("4.  "),
    ti("Aspergillus flavus"), t(" \u2014 Colony appeared yellow-green to olive or blue-green on the surface, with a pale reverse. This colour range is typical of "),
    ti("A. flavus"), t(" on SDA, though it can overlap with other species in Section "),
    ti("Flavi"), t(". The identification is considered tentative and would benefit from microscopic confirmation (examination of conidial head structure and vesicle diameter) or molecular methods in future work.")
  ], { indent: { left: 360 }, spacing: { after: 120, line: 250 } }),

  bp([
    t("5.  "),
    ti("Trichoderma erinaceum"), t(" (white) \u2014 Fast-growing, white, cottony colony that spread rapidly across the plate surface. "),
    ti("Trichoderma"), t(" species are easily recognised by their rapid growth rate and dense, tufted mycelium. The species-level identification was based on the colony\u2019s white colouration and growth pattern, though "),
    ti("Trichoderma"), t(" species can be difficult to differentiate from one another on macroscopic features alone.")
  ], { indent: { left: 360 }, spacing: { after: 120, line: 250 } }),

  bp([
    t("6.  "),
    ti("Aspergillus melleus"), t(" \u2014 Colony with a yellowish-brown to tan appearance, consistent with published descriptions of "),
    ti("A. melleus"), t(" on standard media. This species belongs to Section "),
    ti("Circumdati"), t(" and is less commonly reported in indoor air surveys than "),
    ti("A. niger"), t(" or "),
    ti("A. flavus"), t(", making it a noteworthy finding (Houbraken et al., 2020).")
  ], { indent: { left: 360 }, spacing: { after: 120, line: 250 } }),

  bp([
    t("7.  "),
    ti("Aspergillus japonicus"), t(" (black and white) \u2014 Colony showed a distinctive two-tone appearance, with black and white sectors. Bicoloured or sectoring colonies can arise when different strains or genotypes are present on the same plate, or as a response to local nutrient or pH gradients on the agar surface. The identification was made by comparing the colony\u2019s features with published descriptions of "),
    ti("A. japonicus"), t(" (Klich, 2002).")
  ], { indent: { left: 360 }, spacing: { after: 120, line: 250 } }),

  bp([
    t("It should be noted that all identifications in this study were based on macroscopic colony characteristics only. Microscopic examination of conidial structures was not performed, and no molecular techniques (PCR, sequencing) were used. Species-level identification of "),
    ti("Aspergillus"), t(" and "),
    ti("Penicillium"), t(" from colony morphology alone carries a degree of uncertainty, and the identifications presented here should be regarded as preliminary. This is discussed further in the limitations section of Chapter 5.")
  ]),
];

// ==================== REFERENCES ====================
const refs = [
  'Awad, A.H.A. and Mawla, H.A. (2012). Sedimentation with the Omeliansky Formula as an Accepted Technique for Quantifying Airborne Fungi. Polish Journal of Environmental Studies, 21(5), 1317\u20131320.',
  'Barnett, H.L. and Hunter, B.B. (1998). Illustrated Genera of Imperfect Fungi. 4th edition. APS Press, St. Paul, MN.',
  'Black, W.D. (2020). A comparison of several media types and basic techniques used to assess outdoor airborne fungi in Melbourne, Australia. PLoS ONE, 15(12), e0238901.',
  'Ghazanfari, M., Yazdani Charati, J., Keikha, N., Kholoujini, M., et al. and Hedayati, M.T. (2023). Indoor environment assessment of special wards of educational hospitals for the detection of fungal contamination sources: A multi-center study (2019\u20132021). Current Medical Mycology, 9(3).',
  'Guinea, J., Pel\u00E1ez, T., Alhambra, A. and Bouza, E. (2005). Evaluation of Czapeck agar and Sabouraud dextrose agar for the culture of airborne Aspergillus conidia. Medical Mycology, 43(5), 477\u2013481.',
  'Hocking, A.D. and Pitt, J.I. (1980). Dichloran-glycerol medium for enumeration of xerophilic fungi from low-moisture foods. Applied and Environmental Microbiology, 39(2), 488\u2013492.',
  'Houbraken, J., Frisvad, J.C., Varga, J. and Samson, R.A. (2020). Classification of Aspergillus, Penicillium, Talaromyces and related genera (Eurotiales). Studies in Mycology, 95, 5\u2013169.',
  'Kallawicha, K., Wu, P.-C., Lung, S.-C.C., Lin, Y.-C., Chou, C.-H., Wang, Y.-F. and Su, H.-J. (2017). Ambient fungal spore concentration in a subtropical metropolis: Temporal distributions and meteorological determinants. Aerosol and Air Quality Research, 17, 2051\u20132063.',
  'Klich, M.A. (2002). Identification of Common Aspergillus Species. Centraalbureau voor Schimmelcultures, Utrecht, The Netherlands.',
  'Madukasi, I., Ezejiofor, A.N., Nwaoguikpe, R.N. and Okolo, B.N. (2021). Microbiological indoor air quality within a tertiary institution in south-east, Nigeria. International Journal of Multi-disciplinary Academic Studies, 9(1), 1\u201314.',
  'Pasquarella, C., Pitzurra, O. and Savino, A. (2000). The Index of Microbial Air Contamination. Journal of Hospital Infection, 46, 241\u2013256.',
  'Pitt, J.I. and Hocking, A.D. (2009). Fungi and Food Spoilage. 3rd edition. Springer, New York.',
  'Tomazin, R. and Matos, T. (2024). Mycological Methods for Routine Air Sampling and Interpretation of Results in Operating Theaters. Diagnostics, 14(3), 288.',
  'Viani, I., Colucci, M.E., Veronesi, L. et al. (2020). Passive air sampling: the use of the index of microbial air contamination. Acta Bio Medica: Atenei Parmensis, 91(3-S), 92\u2013105.',
  'Visagie, C.M., Houbraken, J., Frisvad, J.C., Hong, S.-B., Klaassen, C.H.W., Perrone, G., Seifert, K.A., Varga, J., Yaguchi, T. and Samson, R.A. (2014). Identification and nomenclature of the genus Penicillium. Studies in Mycology, 78, 343\u2013371.',
  'Zhang, X., Zhang, R., Wu, Y., Zheng, M. and Wang, P. (2022). Airborne fungal spore review, new advances and automatisation. Atmosphere, 13(2), 308.',
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
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [new TextRun({ text: "\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501", font: "Calibri", size: 20, color: c.accent })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "CHAPTER THREE", font: "Times New Roman", size: 44, bold: true, color: c.primary })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: "MATERIALS AND METHODS", font: "Times New Roman", size: 36, color: c.secondary })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 600 },
          children: [new TextRun({ text: "\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501\u2501", font: "Calibri", size: 20, color: c.accent })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({
            text: "Assessment of Fungal Microflora in the Indoor Air",
            font: "Times New Roman", size: 28, color: c.primary, italics: true,
          })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({
            text: "of Reception Areas in Three Academic Buildings",
            font: "Times New Roman", size: 28, color: c.primary, italics: true,
          })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [new TextRun({
            text: "(Niger, Volta, Limpopo) in Nile University of Nigeria",
            font: "Times New Roman", size: 28, color: c.primary, italics: true,
          })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: "Nile University of Nigeria", font: "Times New Roman", size: 26, bold: true, color: c.primary })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
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
            alignment: AlignmentType.RIGHT,
            spacing: { after: 0 },
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
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.1  Study Area" })] }),
        ...s31,

        // 3.2
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.2  Media Preparation" })] }),
        ...s32,

        // 3.3
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.3  Sample Collection Procedure" })] }),
        ...s33,

        // 3.4
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.4  Colony Counting and Sub-Culturing" })] }),
        ...s34,

        // 3.5
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "3.5  Macroscopic Identification" })] }),
        ...s35,

        // References
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "References" })] }),
        ...refs.map(ref =>
          bp([t(ref)], { indent: { left: 720, hanging: 720 }, spacing: { after: 100, line: 250 } })
        ),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/z/my-project/download/Chapter_Three_Materials_and_Methods.docx", buffer);
  console.log("Done.");
});
