const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Header, Footer,
  AlignmentType, HeadingLevel, PageBreak, PageNumber, BorderStyle
} = require("docx");

// Academic color palette
const c = {
  primary: "1B2A4A",    // Navy
  body: "1E293B",       // Dark slate
  secondary: "4A5568",  // Muted gray
  accent: "8B4513",     // Burgundy-brown
};

// Helper: body paragraph (left-aligned, English, 1.3x line spacing)
function bodyPara(runs, opts = {}) {
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

// ===================== CONTENT =====================

const backgroundParas = [
  // Opening — start with a broad but concrete point, not a sweeping global statement
  bodyPara([
    t("Most people spend a substantial portion of their day indoors — in homes, offices, classrooms, and waiting areas — often assuming that the air they breathe there is clean. In reality, indoor air is a complex mixture of particles, microbes, volatile organic compounds, and other agents that can affect human health in ways we are still working to fully understand. Among the biological components of indoor air, fungi occupy a particularly important place because of their ubiquity, their resilience, and the range of health problems they are known to cause or worsen (Burge, 1990). Fungal spores are present nearly everywhere in the outdoor environment, and they find their way into buildings through open doors and windows, ventilation systems, and on the clothing and bodies of occupants. Once inside, they settle on surfaces, colonise damp or poorly maintained areas, and in many cases reproduce, releasing still more spores into the enclosed air.")
  ]),

  bodyPara([
    t("The World Health Organization has recognised dampness and mould as a risk to health since at least 2009, when it published guidelines linking indoor fungal growth to respiratory symptoms, allergic reactions, and in some cases more serious infections, especially among immunocompromised individuals (WHO, 2009). Khan and Karuppayil (2012), in a widely cited review, described indoor environments as "
    ), ti("potential reservoirs"), t(" of fungal pollution and noted that species commonly found indoors — including members of the genera "),
    ti("Aspergillus"), t(", "), ti("Penicillium"), t(", "), ti("Cladosporium"), t(", and "), ti("Alternaria"), t(" — have well-documented associations with allergic rhinitis, asthma exacerbation, and hypersensitivity pneumonitis. Nevalainen, Täubel, and Hyvärinen (2015) extended this view, arguing that fungi and their secondary metabolites (mycotoxins and volatile organic compounds) are "),
    ti("companions and contaminants"), t(" of indoor spaces — present whether we notice them or not, and capable of influencing health even at relatively low concentrations.")
  ]),

  bodyPara([
    t("Educational buildings deserve special attention in this conversation. Students and staff spend hours each day in lecture halls, laboratories, libraries, and reception areas, often in buildings that were not originally designed with indoor air quality as a priority. The situation in tropical and subtropical regions is especially concerning. Warm temperatures, high relative humidity, and seasonal rainfall create conditions that favour fungal growth on building materials and in the air (Al Hallak et al., 2023). Kallawicha et al. (2017) demonstrated this clearly in their study of ambient fungal spore concentrations in Taipei, a subtropical city, where they found that temperature and relative humidity were the strongest meteorological predictors of airborne fungal load. Similarly, Chin et al. (2020) reported significant associations between indoor fungal exposure and asthma among junior high school students in Johor Bahru, Malaysia — another tropical setting — underscoring the health stakes for young people in these environments.")
  ]),

  bodyPara([
    t("In Nigeria specifically, research on indoor aeromycology has grown but remains limited. Odebode et al. (2020) surveyed airborne fungi across five locations in Lagos State between 2014 and 2016 and found that "),
    ti("Aspergillus"), t(" and "),
    ti("Penicillium"), t(" species dominated the indoor air profiles, a finding consistent with studies from other humid tropical cities. Madukasi et al. (2021) assessed microbiological air quality in lecture halls, laboratories, and offices at a tertiary institution in south-eastern Nigeria and reported fungal counts that exceeded recommended thresholds in several locations. Eze et al. (2021) likewise documented elevated fungal aerosol loads in crowded indoor spaces in Port Harcourt, including schools and markets. These studies, taken together, suggest that fungal contamination of indoor air is a real and measurable problem in Nigerian educational settings, though the body of evidence is still far from comprehensive.")
  ]),

  bodyPara([
    t("This is the context in which the present study sits. Nile University of Nigeria, located in the nation's capital territory of Abuja, operates several academic buildings that serve hundreds of students and staff each day. Among these are the Niger, Volta, and Limpopo buildings — three named halls that house reception areas, lecture rooms, and administrative offices. Like many institutions in the region, the university has expanded rapidly, and questions about the quality of the indoor environment in these buildings have not, to date, been systematically addressed. This chapter sets out the rationale for investigating the fungal microflora in the reception areas of these three buildings, the questions the study seeks to answer, and the contribution it aims to make to the broader literature on indoor air quality in tropical educational institutions.")
  ]),
];

const problemParas = [
  bodyPara([
    t("Despite the growing body of international literature on indoor fungal contamination, there is a persistent gap when it comes to data from sub-Saharan African educational institutions. Much of what we know comes from studies conducted in Europe, North America, and East Asia (Wu et al., 2020; Fan et al., 2021), where climate conditions, building materials, and maintenance practices differ substantially from those found in West Africa. This matters because the factors that drive indoor fungal growth — humidity, temperature, ventilation rates, occupancy density — are not uniform across geographies. A building in Beijing, for example, faces a very different set of indoor environmental pressures than a building in Abuja, and findings from one context cannot simply be transplanted to the other (Lu et al., 2022).")
  ]),

  bodyPara([
    t("Within Nigeria, the few studies that do exist tend to focus on hospitals, residential homes, or outdoor ambient air (Odebode et al., 2020; Eze et al., 2021). Reception areas in university buildings have received comparatively little attention, even though these spaces are often high-traffic zones where students, staff, and visitors congregate — sometimes for extended periods — while waiting for appointments, collecting documents, or attending to administrative matters. The design of many Nigerian university buildings, with open corridors, limited mechanical ventilation, and variable maintenance standards, means that reception areas can behave quite differently from classrooms or laboratories in terms of air exchange and fungal spore dynamics. We simply do not have good baseline data for these spaces.")
  ]),

  bodyPara([
    t("There is also the matter of health risk perception. Students and staff at institutions like Nile University may be routinely exposed to airborne fungal spores without any awareness of the fact, and without any institutional monitoring in place to track exposure levels over time. Mousavi et al. (2016) highlighted the particular concern posed by "),
    ti("Aspergillus"), t(" species in indoor environments, noting their capacity to cause invasive aspergillosis in vulnerable individuals. While healthy adults may tolerate moderate spore exposures without obvious symptoms, the same cannot be said for individuals with asthma, chronic respiratory conditions, or weakened immune systems — populations that are present on any university campus. Without data, there is no basis for risk assessment, and without risk assessment, there is no impetus for improvement.")
  ]),

  bodyPara([
    t("This study therefore addresses a straightforward but important problem: we do not currently know what fungal species are present in the indoor air of the reception areas in the Niger, Volta, and Limpopo buildings at Nile University of Nigeria, nor do we know at what concentrations they occur. That gap in knowledge makes it difficult for the university's administration, and for the broader academic community, to make evidence-based decisions about indoor environmental management.")
  ]),
];

const justificationParas = [
  bodyPara([
    t("A handful of Nigerian studies have touched on indoor air quality in educational settings — Madukasi et al. (2021) in the south-east, Eze et al. (2021) in Port Harcourt, and others — but these were conducted at different institutions, in different climatic zones, and often with different sampling methodologies. Abuja sits in Nigeria's middle belt, characterised by a mix of Sudan savanna vegetation, distinct wet and dry seasons, and an altitude of around 840 metres above sea level. These geographic and climatic features distinguish it from the more southern and coastal cities where most Nigerian aeromycology research has taken place. A study conducted here would add a genuinely new data point to the national literature.")
  ]),

  bodyPara([
    t("The choice of reception areas as the sampling focus is deliberate. While classrooms and laboratories are important, they tend to be occupied in structured, predictable ways — a lecture lasts two hours, a lab session three. Reception areas, by contrast, have irregular, sometimes continuous foot traffic throughout the working day. Doors are opened and closed frequently. People arrive from outside carrying spores on their clothing. The air exchange patterns in these spaces are less well understood, and yet they may represent points of higher or more variable exposure than more controlled indoor environments. Studying reception areas therefore fills a niche that existing campus-based air quality surveys have largely overlooked.")
  ]),

  bodyPara([
    t("Methodologically, this study adopts the passive sedimentation (settle plate) technique for fungal sampling. This is a widely used, low-cost approach that has been employed in aeromycological surveys around the world for decades (Zhang et al., 2022). While it does not capture the total airborne fungal load as precisely as active air sampling, it provides a reliable indication of the culturable fungi that settle out of the air over a given period — which is, arguably, the fraction most relevant to human exposure, since settled spores can later be re-aerosolised through occupant activity, cleaning, or draughts. Yuan et al. (2022) used a combination of active and passive sampling in their survey of indoor and outdoor airborne fungi at Tianjin University in China and found that the two approaches yielded broadly comparable species profiles, lending support to the continued use of settle plates in settings where active samplers are not available.")
  ]),

  bodyPara([
    t("Taken together, the geographic specificity (Abuja), the institutional context (Nile University), the focus on reception areas, and the use of morphological identification methods make this study a meaningful addition to the evidence base. It is not an exhaustive survey of all indoor spaces on campus, but it is a focused, reproducible assessment that can serve as a starting point for more extensive monitoring in the future.")
  ]),
];

const aimObjParas = [
  bodyPara([
    tb("Aim of the Study"), t("")
  ], { spacing: { after: 80, line: 250 } }),
  bodyPara([
    t("The overarching aim of this study is to assess the fungal microflora present in the indoor air of the reception areas in three academic buildings — Niger, Volta, and Limpopo — at Nile University of Nigeria, Abuja.")
  ]),

  new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { after: 80, line: 250 },
    children: [tb("Specific Objectives")],
  }),

  bodyPara([
    t("The specific objectives are to:")
  ]),
  bodyPara([t("1.  Isolate and identify the fungal species present in the indoor air of the reception areas of the three buildings using cultural and morphological methods.")], { indent: { left: 360 } }),
  bodyPara([t("2.  Determine the frequency of occurrence and relative abundance of each fungal species across the three sampling locations.")], { indent: { left: 360 } }),
  bodyPara([t("3.  Compare the fungal diversity profiles among the Niger, Volta, and Limpopo building reception areas.")], { indent: { left: 360 } }),
  bodyPara([t("4.  Assess whether the observed fungal concentrations and species profiles fall within ranges considered acceptable by established indoor air quality guidelines.")], { indent: { left: 360 } }),
];

const significanceParas = [
  bodyPara([
    t("This study contributes to the literature in several practical ways. First, it provides baseline data on indoor fungal contamination at a Nigerian university where no such assessment has previously been carried out. Baseline data is, by its nature, unglamorous — but it is essential. Without knowing what is currently present in the air, it is impossible to track changes over time, evaluate the effectiveness of maintenance interventions, or respond meaningfully to health complaints from building occupants.")
  ]),

  bodyPara([
    t("Second, the findings may inform the university's facilities management decisions. If the study identifies elevated fungal loads or the presence of species with known pathogenic potential (for example, "),
    ti("Aspergillus fumigatus"), t(" or "), ti("Aspergillus flavus"), t("), that information could prompt targeted actions — improved ventilation, moisture control, periodic cleaning protocols, or building material upgrades. Anaissie et al. (2002) demonstrated, in a landmark three-year prospective study in a hospital setting, that pathogenic "),
    ti("Aspergillus"), t(" species can persist in indoor water systems and contribute to nosocomial infections. While a university is not a hospital, the principle is the same: knowing what is there is the first step toward managing it.")
  ]),

  bodyPara([
    t("Third, the study adds to the body of knowledge on indoor air quality in tropical educational buildings more broadly. As Wu et al. (2020) noted in their systematic review of indoor microbial exposures in China, there is a global need for more geographically diverse studies to move beyond the current over-reliance on data from temperate, high-income countries. Nigerian institutions have an important role to play in filling that gap.")
  ]),

  bodyPara([
    t("Finally, from a public health and sustainable development perspective, this study aligns with the Sustainable Development Goals — specifically SDG 3 (Good Health and Well-being) and SDG 4 (Quality Education). Students and staff have a right to learn and work in environments that do not pose unnecessary health risks, and generating the data needed to ensure that is a worthwhile endeavour in itself.")
  ]),
];

const scopeParas = [
  bodyPara([
    t("The study is confined to the reception areas of three academic buildings on the main campus of Nile University of Nigeria in Abuja: the Niger, Volta, and Limpopo buildings. These three buildings were selected because they represent distinct, frequently used points of entry and congregation within the university's academic zone, and because they are accessible for repeated sampling visits within the available timeframe.")
  ]),

  bodyPara([
    t("Sampling is limited to culturable airborne fungi captured using the passive sedimentation (settle plate) method. Fungal identification is carried out through macroscopic and microscopic morphological examination, supported by standard taxonomic keys. The study does not employ molecular identification techniques (such as PCR or DNA sequencing), nor does it attempt to quantify mycotoxin levels or assess other indoor air quality parameters (bacterial load, particulate matter, carbon dioxide concentration, or temperature and humidity readings). These are acknowledged limitations that could be addressed in future work.")
  ]),

  bodyPara([
    t("The temporal scope of the study covers a single sampling period during the rainy season, when fungal spore loads in Abuja are expected to be at or near their peak. While a cross-sectional study of this kind cannot capture seasonal variation, it provides a useful snapshot of worst-case conditions — the period during which indoor fungal contamination is most likely to be elevated. The study does not attempt to correlate fungal findings with health outcomes among building occupants, as that would require a different (and considerably more complex) study design involving clinical data and participant consent.")
  ]),
];

// ===================== DOCUMENT =====================

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Calibri", size: 22, color: c.body },
      },
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
    // ---- COVER PAGE ----
    {
      properties: {
        page: {
          margin: { top: 0, bottom: 0, left: 0, right: 0 },
          size: { width: 11906, height: 16838 },
        },
        titlePage: true,
      },
      children: [
        // Spacer to push content to vertical center
        new Paragraph({ spacing: { before: 4800 }, children: [t("")] }),
        // Decorative line
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [new TextRun({ text: "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", font: "Calibri", size: 20, color: c.accent })],
        }),
        // Title
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "CHAPTER ONE", font: "Times New Roman", size: 44, bold: true, color: c.primary })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: "INTRODUCTION", font: "Times New Roman", size: 36, color: c.secondary })],
        }),
        // Decorative line
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 600 },
          children: [new TextRun({ text: "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", font: "Calibri", size: 20, color: c.accent })],
        }),
        // Project Title
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({
            text: "Assessment of Fungal Microflora in the Indoor Air",
            font: "Times New Roman", size: 28, color: c.primary, italics: true
          })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({
            text: "of Reception Areas in Three Academic Buildings",
            font: "Times New Roman", size: 28, color: c.primary, italics: true
          })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [new TextRun({
            text: "(Niger, Volta, Limpopo) in Nile University of Nigeria",
            font: "Times New Roman", size: 28, color: c.primary, italics: true
          })],
        }),
        // Institution
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({
            text: "Nile University of Nigeria",
            font: "Times New Roman", size: 26, bold: true, color: c.primary
          })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({
            text: "Abuja, Nigeria",
            font: "Calibri", size: 22, color: c.secondary
          })],
        }),
      ],
    },
    // ---- MAIN CONTENT ----
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
            children: [new TextRun({ text: "Chapter One — Introduction", font: "Calibri", size: 18, color: c.secondary, italics: true })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "— ", font: "Calibri", size: 18, color: c.secondary }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Calibri", size: 18, color: c.secondary }),
              new TextRun({ text: " —", font: "Calibri", size: 18, color: c.secondary }),
            ],
          })],
        }),
      },
      children: [
        // 1.1 Background
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun({ text: "1.1  Background of the Study" })],
        }),
        ...backgroundParas,

        // 1.2 Statement of Problem
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun({ text: "1.2  Statement of the Problem" })],
        }),
        ...problemParas,

        // 1.3 Justification
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun({ text: "1.3  Justification of the Study" })],
        }),
        ...justificationParas,

        // 1.4 Aim and Objectives
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun({ text: "1.4  Aim and Objectives of the Study" })],
        }),
        ...aimObjParas,

        // 1.5 Significance
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun({ text: "1.5  Significance of the Study" })],
        }),
        ...significanceParas,

        // 1.6 Scope
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun({ text: "1.6  Scope of the Study" })],
        }),
        ...scopeParas,

        // References heading
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun({ text: "References" })],
        }),
        // Reference list
        ...[
          'Al Hallak, M., Verdier, T., Bertron, A., Roques, C., & Bailly, J.-D. (2023). Fungal contamination of building materials and the aerosolization of particles and toxins in indoor air and their associated risks to health: A review. Toxins, 15(3), 175.',
          'Anaissie, E. J., Stratton, S. L., Dignani, M. C., Lee, C. K., Summerbell, R. C., Rex, J. H., Monson, T. P., & Walsh, T. J. (2002). Pathogenic Aspergillus species recovered from a hospital water system: A 3-year prospective study. Clinical Infectious Diseases, 34(6), 780–789.',
          'Burge, H. A. (1990). Bioaerosols: Prevalence and health effects in the indoor environment. Journal of Allergy and Clinical Immunology, 86(5), 687–701.',
          'Chin, J. Y. W., Zheng, M., Kuo, H.-W., Chang, C.-C., Lee, C.-T., Sahu, S. K., ... & Wu, C.-F. (2020). Indoor microbiome, environmental characteristics and asthma among junior high school students in Johor Bahru, Malaysia. Environment International, 138, 105664.',
          'Eze, E. I., Igwe, V. N., Nwafor, O. I., & Okpokwasili, G. C. (2021). Incidence of fungal aerosols from selected crowded places in Port Harcourt, Nigeria. Asian Journal of Atmospheric Environment, 15(3), 2021036.',
          'Fan, G., Zhang, R., Yang, X., Wang, P., Zhou, X., & Zhang, X. (2021). Investigation of fungal contamination in urban houses with children in six major Chinese cities: Genus and concentration characteristics. Building and Environment, 205, 108229.',
          'Kallawicha, K., Wu, P.-C., Lung, S.-C. C., Lin, Y.-C., Chou, C.-H., Wang, Y.-F., & Su, H.-J. (2017). Ambient fungal spore concentration in a subtropical metropolis: Temporal distributions and meteorological determinants. Aerosol and Air Quality Research, 17, 2051–2063.',
          'Khan, A. A. H., & Mohan Karuppayil, S. (2012). Fungal pollution of indoor environments and its management. Saudi Journal of Biological Sciences, 19(4), 405–426.',
          'Lu, Y., Wang, X., de Souza Almeida, L. C., & Pecoraro, L. (2022). Environmental factors affecting diversity, structure, and temporal variation of airborne fungal communities in a research and teaching building of Tianjin University, China. Journal of Fungi, 8(5), 431.',
          'Madukasi, I., Ezejiofor, A. N., Nwaoguikpe, R. N., & Okolo, B. N. (2021). Microbiological indoor air quality within a tertiary institution in south-east, Nigeria. International Journal of Multi-disciplinary Academic Studies, 9(1), 1–14.',
          'Mousavi, B., Hedayati, M. T., Hedayati, N., Ilkit, M., & Syedmousavi, S. (2016). Aspergillus species in indoor environments and their possible occupational and public health hazards. Current Medical Mycology, 2(1), 36–42.',
          'Nevalainen, A., Täubel, M., & Hyvärinen, A. (2015). Indoor fungi: Companions and contaminants. Indoor Air, 25(2), 125–156.',
          'Odebode, A. C., Adekunle, A. A., Stajich, J. E., & Adeonipekun, P. A. (2020). Airborne fungal spores distribution in various locations in Lagos, Nigeria. Environmental Monitoring and Assessment, 192(2), 87.',
          'World Health Organization. (2009). WHO guidelines for indoor air quality: Dampness and mould. WHO Regional Office for Europe, Copenhagen.',
          'Wu, Y., Zhang, R., Cai, J., Zhang, Y., Zhao, S., Fan, G., ... & Wang, P. (2020). Indoor exposure levels of bacteria and fungi in residences, schools, and offices in China: A systematic review. Indoor Air, 30(6), 1147–1165.',
          'Yuan, C., Wang, X., & Pecoraro, L. (2022). Environmental factors shaping the diversity and spatial-temporal distribution of indoor and outdoor culturable airborne fungal communities in Tianjin University campus, Tianjin, China. Frontiers in Microbiology, 13, 928921.',
          'Zhang, X., Zhang, R., Wu, Y., Zheng, M., & Wang, P. (2022). Airborne fungal spore review, new advances and automatisation. Atmosphere, 13(2), 308.',
        ].map(ref =>
          bodyPara([t(ref)], { indent: { left: 720, hanging: 720 }, spacing: { after: 100, line: 250 } })
        ),
      ],
    },
  ],
});

// Generate file
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/z/my-project/download/Chapter_One_Fungal_Microflora_Nile_University_Revised.docx", buffer);
  console.log("Document saved successfully.");
});
