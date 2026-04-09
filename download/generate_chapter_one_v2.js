const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel,
  Header, Footer, PageNumber, PageBreak, BorderStyle, TabStopType, TabStopPosition,
  LevelFormat
} = require("docx");

const colors = {
  primary: "1A1F16",
  body: "2D3329",
  secondary: "4A5548",
  accent: "94A3B8",
};

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Times New Roman", size: 24, color: colors.body }
      }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, color: colors.primary, font: "Times New Roman" },
        paragraph: { spacing: { before: 480, after: 240, line: 276 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, color: colors.primary, font: "Times New Roman" },
        paragraph: { spacing: { before: 360, after: 180, line: 276 }, outlineLevel: 1 }
      }
    ]
  },
  numbering: {
    config: [
      {
        reference: "objectives-list",
        levels: [{
          level: 0, format: LevelFormat.LOWER_ROMAN, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  sections: [
    // COVER PAGE
    {
      properties: {
        page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, size: { width: 11906, height: 16838 } },
        titlePage: true
      },
      headers: {
        default: new Header({ children: [new Paragraph({ children: [] })] }),
        first: new Header({ children: [new Paragraph({ children: [] })] })
      },
      footers: {
        default: new Header({ children: [new Paragraph({ children: [] })] }),
        first: new Header({ children: [new Paragraph({ children: [] })] })
      },
      children: [
        new Paragraph({ spacing: { before: 2400 }, alignment: AlignmentType.CENTER, children: [] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "ASSESSMENT OF FUNGAL MICROFLORA IN THE", font: "Times New Roman", size: 32, bold: true, color: colors.primary })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "INDOOR AIR OF RECEPTION AREAS IN THREE", font: "Times New Roman", size: 32, bold: true, color: colors.primary })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "ACADEMIC BUILDINGS (NIGER, VOLTA, LIMPOPO)", font: "Times New Roman", size: 32, bold: true, color: colors.primary })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: "IN NILE UNIVERSITY OF NIGERIA", font: "Times New Roman", size: 32, bold: true, color: colors.primary })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: "____________________________", font: "Times New Roman", size: 24, color: colors.accent })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "CHAPTER ONE", font: "Times New Roman", size: 28, bold: true, color: colors.secondary })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: "INTRODUCTION", font: "Times New Roman", size: 28, bold: true, color: colors.secondary })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "A Project Submitted to the Department of Microbiology,", font: "Times New Roman", size: 22, color: colors.secondary })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "Faculty of Sciences, Nile University of Nigeria", font: "Times New Roman", size: 22, color: colors.secondary })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "In Partial Fulfillment of the Requirement for the Award of", font: "Times New Roman", size: 22, color: colors.secondary })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 }, children: [new TextRun({ text: "the Bachelor of Science Degree in Microbiology (B.Sc.)", font: "Times New Roman", size: 22, color: colors.secondary })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "April, 2026", font: "Times New Roman", size: 24, color: colors.secondary })] })
      ]
    },
    // MAIN CONTENT
    {
      properties: {
        page: { margin: { top: 1800, right: 1440, bottom: 1440, left: 1440 }, size: { width: 11906, height: 16838 }, pageNumbers: { start: 1 } }
      },
      headers: {
        default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0 }, children: [new TextRun({ text: "Chapter One: Introduction", font: "Times New Roman", size: 18, italics: true, color: colors.accent })] })] })
      },
      footers: {
        default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", font: "Times New Roman", size: 18, color: colors.accent }), new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 18, color: colors.accent }), new TextRun({ text: " of ", font: "Times New Roman", size: 18, color: colors.accent }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Times New Roman", size: 18, color: colors.accent })] })] })
      },
      children: [
        new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, spacing: { before: 0, after: 360 }, children: [new TextRun({ text: "CHAPTER ONE", font: "Times New Roman", size: 36, bold: true, color: colors.primary })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 480 }, children: [new TextRun({ text: "INTRODUCTION", font: "Times New Roman", size: 32, bold: true, color: colors.primary })] }),

        // 1.1 BACKGROUND OF THE STUDY
        new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 360, after: 240 }, children: [new TextRun({ text: "1.1  Background of the Study", font: "Times New Roman", size: 28, bold: true, color: colors.primary })] }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "Fungi represent one of the most diverse and ecologically significant groups of microorganisms on Earth, playing indispensable roles in nutrient cycling, organic matter decomposition, and the maintenance of balanced ecosystems. They exist in virtually every habitat, from soil and water to the atmosphere, where they disperse primarily through the release of microscopic spores or fragmented hyphae that are easily carried by air currents. Airborne particles which are biological in origin are referred to as bioaerosols, and these include fungal spores, pollens, bacteria, mites, and dead tissue fragments (Kour et al., 2025). In outdoor environments, airborne fungi contribute to natural processes such as the breakdown of plant litter and the recycling of essential nutrients. However, when these fungal propagules enter indoor spaces, they can accumulate to levels that pose significant risks to human health, building infrastructure, and the integrity of stored materials. The study of fungal communities in indoor air, known as aeromycology, has therefore emerged as a critical area of environmental microbiology and public health research.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "Indoor air quality has become a subject of growing global concern as populations worldwide spend increasing proportions of their time within enclosed spaces, including homes, offices, hospitals, and educational institutions. The World Health Organization (WHO) has estimated that individuals in urban areas may spend up to 90% of their daily lives indoors, making the quality of indoor air a major determinant of overall health and well-being (WHO, 2009). Among the various biological contaminants that compromise indoor air quality, fungal spores are particularly noteworthy due to their ubiquity, resilience, and capacity to cause a wide spectrum of adverse health effects. Common indoor fungal genera such as Aspergillus, Penicillium, Cladosporium, Alternaria, and Mucor have been repeatedly identified as dominant components of indoor mycoflora across diverse geographic and climatic settings (Shelton et al., 2002; Verma et al., 2012). These fungi produce vast quantities of spores that remain suspended in air for extended periods, settle on surfaces, and colonise available substrates under favourable conditions of temperature, humidity, and nutrient availability.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "The health implications of exposure to indoor fungal spores are extensive and well-documented in the scientific literature. Inhalation of fungal spores and their fragmented hyphal components can trigger allergic reactions, including allergic rhinitis, conjunctivitis, and dermatological hypersensitivity responses. More severe consequences include the exacerbation of asthma, the development of hypersensitivity pneumonitis, and invasive fungal infections, particularly among immunocompromised individuals such as patients undergoing chemotherapy, organ transplant recipients, and those living with HIV/AIDS (Al-Shaarani and Pecoraro, 2024). Filamentous fungi from various environmental sources have been associated with allergic rhinitis, asthma, tonsillitis, adenoid hypertrophy, and impaired lung function, as demonstrated in a recent study on potential respiratory hazards of fungal exposure (Ozkara et al., 2025). Certain fungal species, notably members of the genera Aspergillus and Fusarium, are also known to produce mycotoxins, which are toxic secondary metabolites that can cause acute and chronic health effects when inhaled or ingested. Hallak et al. (2023) provided a comprehensive review of the fungal contamination of building materials and the aerosolisation of particles and toxins in indoor air, emphasising the associated health risks.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "Educational institutions represent a particularly important category of indoor environments for fungal air quality assessment. University campuses typically host large and diverse populations of students, faculty members, administrative staff, and visitors who spend considerable periods within classrooms, lecture halls, libraries, offices, and reception areas. These spaces are characterised by high occupancy density, frequent movement of people, and variable environmental conditions, all of which contribute to the accumulation and dispersal of airborne fungal spores. Ilie\u0219 et al. (2025) conducted an integrated analysis of indoor air quality and fungal microbiota in educational heritage buildings, revealing that microclimate conditions, building occupancy patterns, and ventilation efficiency significantly influenced both fungal diversity and concentration. Reception areas in academic buildings are especially noteworthy in this regard, as they serve as primary points of entry and congregation, where visitors, students, and staff continuously introduce outdoor fungal spores while simultaneously disturbing settled dust that may harbour fungal propagules.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "In the Nigerian context, the assessment of indoor fungal microflora carries particular urgency due to the prevailing tropical climate, which is characterised by elevated temperatures and high relative humidity throughout much of the year. These environmental conditions are highly conducive to fungal growth, sporulation, and spore dispersal, creating conditions that favour the proliferation of diverse fungal species within enclosed buildings. Several studies conducted within Nigerian tertiary institutions have confirmed the presence of significant fungal loads in indoor air across various building types. Fakunle et al. (2022) investigated associations between indoor bacterial and fungal aerosols and lower respiratory tract infections among under-five children in Ibadan, Nigeria, providing compelling evidence that indoor fungal exposure is a significant predictor of respiratory illness in the Nigerian setting. Eghomwanre and Imarhiagbe (2025) assessed the associations between indoor bioaerosol concentrations and the prevalence of childhood asthma in Benin City, Nigeria, further highlighting the public health burden of indoor fungal contamination in Nigerian urban centres. More recently, the microbial air quality of a tertiary hospital in Makurdi, North-Central Nigeria, was assessed, revealing substantial airborne bacterial and fungal loads across multiple indoor environments (Ogunlade et al., 2026). These Nigerian studies collectively underscore the pervasiveness of fungal contamination in indoor environments and the corresponding health risks faced by building occupants.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "Research conducted in other tropical and subtropical regions has corroborated the Nigerian findings and broadened the understanding of indoor fungal dynamics. Ahmad et al. (2025) assessed fungi in indoor and outdoor environments of four selected hospitals in Peninsular Malaysia, isolating 68 fungal genera and identifying Aspergillus spp. (21.1%) and Penicillium spp. (20.1%) as the most common genera in indoor air samples. Dang et al. (2025) investigated airborne fungal microbiota in indoor and outdoor classrooms of schools in Ho Chi Minh City, Vietnam, reporting significant variations in fungal concentration and composition between indoor and outdoor environments. Tadesse et al. (2024) investigated microbial contamination of indoor air, environmental surfaces, and medical equipment in a hospital in southwestern Ethiopia, documenting the interplay between environmental factors and microbial load in tropical healthcare settings. Dela Rosa et al. (2025) quantified indoor fungal aerosol levels in slum households located in tropical humid regions, concluding that indoor fungal load was a major determinant of the health risk quotient for residents. Cervantes et al. (2025) published a critical review of fungal contamination in schools, providing a comprehensive assessment of the methodological approaches and health implications of fungal exposure in educational buildings worldwide. These studies confirm that indoor fungal contamination is a global concern, but one that is particularly acute in tropical regions where climatic conditions favour fungal proliferation.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "Nile University of Nigeria, located in Abuja, the Federal Capital Territory, is a private institution that has experienced significant growth in its student population and physical infrastructure. The university campus comprises several academic buildings, including the Niger, Volta, and Limpopo buildings, which house various faculties and departments. These academic buildings feature reception areas that serve as central points of access and interaction for students, faculty, staff, and visitors. Given the tropical climate of Abuja, characterised by distinct wet and dry seasons with consistently warm temperatures and periodic elevated humidity levels, the indoor environments within these buildings are potentially susceptible to fungal colonisation and spore accumulation. Despite the critical importance of indoor air quality for the health and productivity of the university community, there is currently a paucity of data regarding the fungal microflora present in the indoor air of reception areas within the Niger, Volta, and Limpopo academic buildings. The present study is therefore designed to address this gap by conducting a systematic assessment of the fungal microflora in the indoor air of reception areas in these three academic buildings at Nile University of Nigeria.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        // 1.2 STATEMENT OF THE RESEARCH PROBLEM
        new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 480, after: 240 }, children: [new TextRun({ text: "1.2  Statement of the Research Problem", font: "Times New Roman", size: 28, bold: true, color: colors.primary })] }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "Indoor air quality in educational institutions is a critical determinant of the health, comfort, and academic performance of students, faculty, and staff. Among the biological contaminants that compromise indoor air quality, airborne fungi occupy a position of particular concern due to their capacity to cause allergic reactions, respiratory infections, toxicosis through mycotoxin production, and structural damage to building materials and stored items. Sultan et al. (2025) conducted a rapid review of exposure risks from microbiological hazards in buildings and their implications for human health, highlighting the evidence for human exposure to harmful microorganisms from indoor air and surfaces in built environments. Reception areas in academic buildings represent especially vulnerable indoor spaces because they function as high-traffic zones where large numbers of individuals congregate, facilitating the continuous introduction and resuspension of fungal spores from both outdoor and indoor sources. The design and operational characteristics of these spaces, including the frequency of door and window openings, the density of human occupancy, and the adequacy of ventilation systems, can significantly influence the concentration and composition of airborne fungal communities.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "At Nile University of Nigeria, the Niger, Volta, and Limpopo academic buildings are major hubs of academic and administrative activity. Their reception areas receive a continuous influx of students, lecturers, administrative personnel, and visitors throughout the academic day, creating conditions that are potentially conducive to the accumulation and dispersal of airborne fungal spores. The tropical climate of Abuja, with its elevated ambient temperatures and seasonal spikes in relative humidity during the rainy months, further compounds the risk of fungal proliferation in these enclosed spaces. Despite these recognised risk factors, there has been no prior systematic assessment of the fungal microflora present in the indoor air of the reception areas within these three academic buildings. This absence of data means that the university community may be unknowingly exposed to fungal species that pose health risks, particularly to individuals with pre-existing respiratory conditions such as asthma, allergic rhinitis, or compromised immune systems.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "The lack of baseline data on indoor fungal contamination also impedes the development of targeted intervention strategies for improving air quality within these buildings. Without an understanding of the specific fungal genera and species present, their relative abundance, and the environmental factors driving their proliferation, any efforts to mitigate indoor air contamination would be largely speculative and potentially ineffective. Ili\u0107 et al. (2023) published a comprehensive review of microbial contamination in the indoor built environment, providing an overview of sources, sampling strategies, and analysis methods, and highlighting a critical knowledge gap related to microbiological burden in indoor environments in developing countries. This research problem is further underscored by the findings of previous studies in similar Nigerian educational settings, which have consistently reported elevated fungal loads in indoor air, with the predominance of potentially pathogenic genera such as Aspergillus, Penicillium, and Cladosporium. The present study therefore seeks to fill this critical knowledge gap by providing a comprehensive assessment of the fungal microflora in the reception areas of the Niger, Volta, and Limpopo buildings.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        // 1.3 JUSTIFICATION OF THE STUDY
        new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 480, after: 240 }, children: [new TextRun({ text: "1.3  Justification of the Study", font: "Times New Roman", size: 28, bold: true, color: colors.primary })] }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "The justification for this study is rooted in the fundamental importance of maintaining healthy indoor air quality in educational institutions and the specific vulnerability of reception areas in academic buildings to fungal contamination. Indoor air quality has been recognised by the World Health Organization and other international public health bodies as a major environmental health determinant, with biological contaminants such as fungal spores identified as key contributors to the sick building syndrome, a constellation of health symptoms that are associated with occupancy of certain indoor environments (WHO, 2009). Students, lecturers, and administrative staff who spend prolonged periods in indoor environments with elevated fungal loads are at increased risk of developing or exacerbating respiratory conditions, allergic responses, and other fungal-related health problems. Ilie\u0219 et al. (2025) demonstrated that poor indoor air quality in educational buildings, often resulting from inadequate ventilation and fungal contamination, leads to what is known as Sick Building Syndrome (SBS), which is a term used to describe a range of non-specific health symptoms experienced by building occupants, including headaches, fatigue, respiratory irritation, and difficulty concentrating.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "Reception areas in the Niger, Volta, and Limpopo academic buildings at Nile University of Nigeria represent strategic sampling points for several reasons. Firstly, these areas are characterised by the highest foot traffic among all spaces within the buildings, making them potential hotspots for the introduction and redistribution of fungal spores from both outdoor and indoor sources. Secondly, reception areas are often located near building entrances, where the influence of outdoor air is most pronounced, and where temperature and humidity gradients between the exterior and interior environments may create microclimatic conditions that favour fungal survival and growth. Thirdly, these areas are typically furnished with materials such as carpets, upholstered furniture, and wooden fixtures that can trap dust and organic matter, providing substrates for fungal colonisation. Cervantes et al. (2025) emphasised in their comprehensive review that fungal contamination in schools has potential negative outcomes on both the health and learning ability of students, reinforcing the need for regular environmental monitoring in educational institutions. Despite the clear rationale for monitoring fungal air quality in these spaces, no prior study has specifically targeted the reception areas of these three academic buildings, making the present investigation both novel and timely.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "Furthermore, the findings of this study will provide essential baseline data that can serve multiple practical purposes. For the university administration, the results will offer a scientific basis for evaluating current building maintenance practices, ventilation adequacy, and cleaning protocols, and for implementing targeted improvements where necessary. For the broader academic community, the data will contribute to the limited but growing body of literature on indoor aeromycology in Nigerian tertiary institutions, facilitating comparisons with studies conducted in other universities within Nigeria and across the West African subregion. The study also aligns with the global emphasis on sustainable building management practices, as advocated by contemporary frameworks such as the United Nations Sustainable Development Goals (SDGs), particularly SDG 3 (Good Health and Well-being) and SDG 4 (Quality Education), both of which underscore the importance of providing healthy and safe learning environments.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        // 1.4 AIM AND OBJECTIVES
        new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 480, after: 240 }, children: [new TextRun({ text: "1.4  Aim and Objectives of the Study", font: "Times New Roman", size: 28, bold: true, color: colors.primary })] }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "The aim of this study is to assess the fungal microflora present in the indoor air of reception areas in three academic buildings, namely Niger, Volta, and Limpopo, at Nile University of Nigeria. This overarching aim is pursued through the following specific objectives:",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 160, line: 360 }, indent: { left: 720 }, numbering: { reference: "objectives-list", level: 0 }, children: [new TextRun({ text: "To determine the fungal load (colony-forming units per cubic metre of air) in the indoor air of reception areas in the Niger, Volta, and Limpopo academic buildings.", font: "Times New Roman", size: 24, color: colors.body })] }),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 160, line: 360 }, indent: { left: 720 }, numbering: { reference: "objectives-list", level: 0 }, children: [new TextRun({ text: "To isolate and identify the predominant fungal genera and species present in the indoor air samples collected from the reception areas of the three academic buildings.", font: "Times New Roman", size: 24, color: colors.body })] }),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 160, line: 360 }, indent: { left: 720 }, numbering: { reference: "objectives-list", level: 0 }, children: [new TextRun({ text: "To compare the diversity and abundance of airborne fungal species among the three academic buildings and assess whether statistically significant differences exist between them.", font: "Times New Roman", size: 24, color: colors.body })] }),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 }, indent: { left: 720 }, numbering: { reference: "objectives-list", level: 0 }, children: [new TextRun({ text: "To evaluate the potential health implications of the identified fungal species for occupants of the buildings and provide evidence-based recommendations for improving indoor air quality.", font: "Times New Roman", size: 24, color: colors.body })] }),

        // 1.5 SIGNIFICANCE
        new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 480, after: 240 }, children: [new TextRun({ text: "1.5  Significance of the Study", font: "Times New Roman", size: 28, bold: true, color: colors.primary })] }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "This study holds considerable significance for multiple stakeholders, including the immediate university community, the broader field of environmental microbiology, public health policy, and institutional facility management. From a public health perspective, the identification and quantification of fungal species in the reception areas of the Niger, Volta, and Limpopo buildings will provide critical information about the potential exposure risks faced by students, faculty, staff, and visitors. By documenting the specific fungal genera present and their relative abundance, the study will enable healthcare providers within the university health services to better understand the mycological dimensions of indoor air-related health complaints, such as respiratory distress, allergic reactions, and unexplained flu-like symptoms that may be attributable to fungal exposure. Al-Shaarani and Pecoraro (2024) provided a comprehensive review of pathogenic airborne fungi and bacteria, unveiling their occurrence, sources, and profound human health implications, thereby reinforcing the public health importance of regular indoor air quality monitoring. This knowledge is particularly important for protecting vulnerable subpopulations within the university community, including individuals with asthma, allergic rhinitis, atopic dermatitis, and those with immunocompromised conditions.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "From an institutional management standpoint, the findings of this study will provide Nile University of Nigeria with empirically grounded evidence to guide decisions regarding building maintenance, ventilation improvement, and cleaning protocols within the academic buildings. If the assessment reveals fungal loads that exceed recommended thresholds, or identifies the presence of known pathogenic or toxigenic species, the university administration will be positioned to take proactive corrective measures, such as upgrading ventilation systems, implementing regular air quality monitoring programmes, or commissioning professional building hygiene assessments. These interventions have the potential to reduce occupant health complaints, enhance comfort and satisfaction within the learning environment, and contribute to the overall reputation of the university as an institution that prioritises the health and safety of its community. The study also provides a model that could be replicated across other buildings within the university campus, thereby supporting a comprehensive, campus-wide approach to indoor air quality management.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "From an academic and scientific perspective, this study contributes valuable data to the existing body of knowledge on indoor aeromycology in Nigerian educational institutions. While several studies have examined indoor air quality in Nigerian universities, the specific focus on reception areas of academic buildings at Nile University of Nigeria represents a novel contribution that addresses a previously uninvestigated aspect of the campus environment. The comparative analysis of fungal diversity and abundance across three distinct buildings will generate insights into the influence of building-specific factors, such as occupancy patterns, ventilation design, and maintenance practices, on indoor fungal communities. Kour et al. (2025) reviewed the diversity and health impacts of airborne fungal communities and highlighted the transformative potential of artificial intelligence applications in aeromycology, noting that improved detection and monitoring technologies could significantly enhance the accuracy of indoor fungal assessments. These findings will be of interest to researchers in the fields of environmental microbiology, mycology, indoor air quality science, and building ecology.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "Additionally, the study has implications for the development of indoor air quality guidelines and standards specifically tailored to the Nigerian educational context. Currently, many Nigerian institutions rely on international indoor air quality standards that may not fully account for the unique climatic, architectural, and cultural factors that influence indoor air quality in tropical African settings. The data generated by this study will contribute to the evidence base needed to support the formulation of context-appropriate national guidelines for indoor fungal contamination in educational and other public buildings in Nigeria.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        // 1.6 SCOPE
        new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 480, after: 240 }, children: [new TextRun({ text: "1.6  Scope of the Study", font: "Times New Roman", size: 28, bold: true, color: colors.primary })] }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "This study is focused specifically on the assessment of fungal microflora in the indoor air of reception areas within three academic buildings, namely the Niger, Volta, and Limpopo buildings, at Nile University of Nigeria, Abuja. The scope of the investigation encompasses the collection of air samples from the reception areas of each building using established microbiological sampling techniques, the cultivation of fungal colonies on appropriate culture media, and the identification and characterisation of fungal isolates based on their macroscopic and microscopic morphological features. The study will determine the fungal load at each sampling site, expressed as colony-forming units per plate, and will identify the predominant fungal genera and species present. A comparative analysis of fungal diversity and abundance across the three buildings will be conducted to assess whether significant inter-building differences exist.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { after: 200, line: 360 },
          children: [new TextRun({ text: "The study is limited to the mycological assessment of indoor air and does not extend to the investigation of other biological air contaminants such as bacteria, viruses, or pollen. Furthermore, the study is confined to the reception areas of the specified buildings and does not include other indoor spaces such as classrooms, laboratories, libraries, or offices within the same buildings. The identification of fungal isolates will be based on conventional morphological and microscopic methods, and the study will not employ molecular techniques such as polymerase chain reaction (PCR) or DNA sequencing for species-level identification. Environmental parameters such as temperature and relative humidity may be recorded at the time of sampling to provide contextual information, but a detailed investigation of the relationship between these parameters and fungal load falls outside the primary scope of this study. The temporal scope of the study is limited to a single sampling period, and the findings should therefore be interpreted as representing a snapshot of the fungal microflora at a specific point in time rather than a comprehensive characterisation of year-round indoor fungal dynamics.",
            font: "Times New Roman", size: 24, color: colors.body
          })]
        }),

        // REFERENCES (Chapter One only)
        new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 480, after: 240 }, children: [new TextRun({ text: "References", font: "Times New Roman", size: 28, bold: true, color: colors.primary })] }),

        ...[
          "Ahmad, N.I., Ismail, S.N.S., Mohamed, S.M., Mohamed, R.A. and Lee, H.L. (2025). Fungi assessment in indoor and outdoor environment of four selected hospitals in Peninsular Malaysia. Journal of Fungi. https://doi.org/10.3390/jof11040318",
          "Al-Shaarani, A. and Pecoraro, L. (2024). A review of pathogenic airborne fungi and bacteria: Unveiling occurrence, sources, and profound human health implication. Environmental Research Communications, 12: 121008. https://doi.org/10.1088/2515-7620/ad8d47",
          "Cervantes, R., Pena, P., Riesenberger, B., Rodriguez, M. and Henderson, S. (2025). Critical insights on fungal contamination in schools: A comprehensive review of assessment methods. Frontiers in Public Health, 13: 1557506. https://doi.org/10.3389/fpubh.2025.1557506",
          "Dang, D.Y.N., Pham, T., Nguyen, T.T., Phan, V.T. and Nguyen, H.T. (2025). Airborne fungal microbiota in indoor and outdoor classrooms of schools in Ho Chi Minh City: Concentration and composition. In\u017cynieria Mineralna (Mineral Engineering). https://doi.org/10.29227/IM-2025-01-05",
          "Dela Rosa, J., Jain, A. and Chakraborty, S. (2025). Quantifying indoor fungal aerosol levels in slums: Implications for public health. Annals of Occupational Hygiene. https://doi.org/10.1093/annhyg/meaf011",
          "Eghomwanre, A.F. and Imarhiagbe, E.E. (2025). Associations between indoor bioaerosol concentrations and the prevalence of childhood asthma in Benin City, Nigeria. Journal of Materials and Environmental Sciences, 16(11): 1611-118. https://doi.org/10.26872/jmes.2025.1611118",
          "Fakunle, A.G., Jafta, N., Smit, L.A.M. and Naidoo, R.N. (2022). Indoor bacterial and fungal aerosols as predictors of lower respiratory tract infections among under-five children in Ibadan, Nigeria. BMC Infectious Diseases, 22: 471. https://doi.org/10.1186/s12879-022-07885-4",
          "Hallak, H., Verdier, T., Mohamad, M., Bertron, A., Roques, C. and Bailly, J.D. (2023). Fungal contamination of building materials and the aerosolization of particles and toxins in indoor air and their associated risks to health: A review. Toxins, 15(3): 175. https://doi.org/10.3390/toxins15030175",
          "Ilie\u0219, A., Burt\u0103, O., Al Shomali, M.A., Gherghel, I., Muntean, V. and Covaci, A. (2025). Integrated analysis of indoor air quality and fungal microbiota in educational heritage buildings: Implications for health and sustainability. Sustainability, 17(3): 1091. https://doi.org/10.3390/su17031091",
          "Ili\u0107, D., Radojevi\u0107, D., Stanojevi\u0107, J.P., \u0160ulc, J. and \u0106ur\u010di\u0107, S. (2023). A comprehensive review of microbial contamination in the indoor built environment. Frontiers in Public Health, 11: 1285393. https://doi.org/10.3389/fpubh.2023.1285393",
          "Kour, D., Khan, S.S., Kour, H., Kour, H., Singh, M. and Yadav, A.N. (2025). Airborne fungal communities: Diversity, health impacts, and potential AI applications in aeromycology. Aerobiologia, 3(4): 10. https://doi.org/10.3390/aerobiologia3040010",
          "Ogunlade, A., Oyelami, O.O. and Oladele, T.T. (2026). Microbial air quality assessment of a tertiary hospital in Makurdi, North-Central Nigeria. Science Letters, 14(1). https://doi.org/10.47262/SL/14/1/03",
          "Ozkara, A., Pinar, N.M., D\u00FCzg\u00FCnes, E., B\u0131y\u00FCkkarakas, E. and Palab\u0131y\u0131k, S.S. (2025). Potential respiratory hazards of fungal exposure in the indoor environment. Mikrobiyoloji B\u00fclteni. https://doi.org/10.5578/mb.2025.10523",
          "Shelton, B.G., Kirkland, K.H., Flanders, W.D. and Morris, G.K. (2002). Profiles of airborne fungi in buildings and outdoor environments in the United States. Applied and Environmental Microbiology, 68(4): 1743-1753. https://doi.org/10.1128/AEM.68.4.1743-1753.2002",
          "Sultan, M., Azam, K., Inayatullah, S., Aslam, F., Khan, M.S. and Saeed, M. (2025). Exposure risks from microbiological hazards in buildings and their implications for human health: A rapid review. Atmosphere, 16(11): 1243. https://doi.org/10.3390/atmos16111243",
          "Tadesse, B., Lemma, A., Alemu, A. and Argaw, T. (2024). Investigating microbial contamination of indoor air, environmental surfaces, and medical equipment in a Southwestern Ethiopia hospital. Environmental Health Insights, 18. https://doi.org/10.1177/11786302241259332",
          "Verma, S., Srinivas, B. and Tyagi, R.D. (2012). Fungal pollution of indoor environments and its management. Saudi Journal of Biological Sciences, 19(4): 405-426. https://doi.org/10.1016/j.sjbs.2012.06.002",
          "WHO (2009). WHO Guidelines for Indoor Air Quality: Dampness and Mould. World Health Organization, Copenhagen. ISBN: 9789289041683."
        ].map(ref => new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 120, line: 276 },
          indent: { left: 720, hanging: 720 },
          children: [new TextRun({ text: ref, font: "Times New Roman", size: 22, color: colors.body })]
        }))
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/z/my-project/download/Chapter_One_Fungal_Microflora_Nile_University.docx", buffer);
  console.log("Chapter One DOCX v2 generated successfully!");
});
