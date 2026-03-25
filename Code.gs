/* ------------------ UI, PRONOUN HELPER, GRAMMAR DICTIONARY, & GENERAL REGISTRY ------------------ */
  // Medical Template UI 
    function onOpen() {
    const ui = DocumentApp.getUi();
  
    // Add menu items
    ui.createMenu("Dr. I.E. Osadiaye's Medical Templates")
    .addItem('Open Medical Template', 'showSidebar')
    .addItem('Expand Keywords', 'expandKeywords')
    .addToUi();

    showSidebar();
    }

  // Pronoun Helper
    function getPronoun(sex) {
    return sex === 'female'
    ? { he: 'she', He: 'She', him: 'her', Him: 'Her', his: 'her', His: 'Her' }
    : { he: 'he', He: 'He', him: 'him', Him: 'Him', his: 'his', His: 'His' };
    }

  // Grammar Dictionary
    const GRAMMAR_DICTIONARY = {
    cherry_eye: {
    singular: { gland: "gland", eye: "eye", make: "makes", is: "is", its: "its", this: "this", pro: "it" },
    plural:   { gland: "glands", eye: "eyes", make: "make", is: "are", its: "their", this: "these", pro: "they" }
    },

    patella: {
    singular: { sub: "a luxating patella", bone: "kneecap", v: "slides", is: "is", pos: "its" },
    plural:   { sub: "luxating patellae", bone: "kneecaps", v: "slide", is: "are", pos: "their" }
    },

    wellness: {
    singular: { 
        dog: "dog", dogs: "dog's", Dogs: "Dog", puppy: "puppy", is: "is", has: "has", was: "was", 
        eats: "eats", weighs: "weighs", steals: "steals", begins: "begins", gets: "gets", 
        mother: "mother's", site: "site", shot: "shot", round: "round", 
        he: "he", him: "him", his: "his", them: "them" 
    },
    plural: { 
        dog: "dogs", dogs: "dogs'", Dogs: "Dogs", puppy: "puppies", is: "are", has: "have", was: "were", 
        eats: "eat", weighs: "weigh", steals: "steal", begins: "begin", gets: "get", 
        mother: "mothers'", site: "sites", shot: "shots", round: "rounds", 
        he: "they", him: "them", his: "their", them: "them" 
    }
    }
    };

    // Plural Pronoun Helper
      function getGrammar(system, plurality = 'singular', sex = 'male') {
      const isPlural = (plurality === 'plural');

      // 1. Get the base words from the dictionary
      const words = isPlural ? GRAMMAR_DICTIONARY[system].plural : GRAMMAR_DICTIONARY[system].singular;

      // 2. Generate the pronouns
      let pronouns = {};
      if (isPlural) {
      pronouns = { he: 'they', He: 'They', him: 'them', Him: 'Them', his: 'their', His: 'Their' };
      } else {
      pronouns = (sex === 'female') 
      ? { he: 'she', He: 'She', him: 'her', Him: 'Her', his: 'her', His: 'Her' }
      : { he: 'he', He: 'He', him: 'him', Him: 'Him', his: 'his', His: 'His' };
      }

      // 3. Merge words and pronouns into one "g" object
      return { ...words, ...pronouns };
      }

  // Style Registry
    const STYLE_REGISTRY = {
    bold: { bold: true },
    boldUnderline: { bold: true, underline: true },
    italic: { italic: true },
    green: { bold: true, underline: true, color: '#6aa84f' },
    red: { bold: true, underline: true, color: '#ff0000' },
    doubleSpaced: { lineSpacing: 2.0 },
    title: { alignment: 'center', fontSize: 20, lineSpacing: 2.0 }
    };

  // Table Colour Registry
    const TABLE_COLOR_REGISTRY = {
    red: '#EA9999',
    yellow: '#FFE599',
    green: '#B6D7A8',
    blue: '#A4C2F4',
    purple: '#B4A7D6',
    };
  
/* ------------------ MEDICINE REGISTRY ------------------ */
  // Prescriptions
    let TABLE_ROW_BUFFER = [];
    const MEDICINE_REGISTRY = {

    ADEQUANINJECTION: {
    label: "Adequan",
    instructions: "Injected in your dog’s muscle to improve joint health.",
    class: "Polysulfated glycosaminoglycan",
    sideEffects: "Well tolerated"
    },

    APOQUEL: {
    label: "Apoquel 3.6mg (oclacitinib)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of allergies.",
    class: "Anti-allergy (JAK inhibitor)",
    sideEffects: "Over-suppresses the immune system when given with Zenrelia."
    },

    CARPROFEN: {
    label: "Carprofen 25mg",
    instructions: "Give your dog 1 tablet by mouth every 12 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite. DO NOT USE WITHIN 3 DAYS OF OTHER NSAIDs OR STEROIDS."
    },

    CYTOPOINTINJECTION: {
    label: "Cytopoint injection (lokivetmab)",
    instructions: "Medication injected beneath your dog’s skin to control allergies over the next 4 - 8 weeks.",
    class: "Anti-allergy (monoclonal antibody)",
    sideEffects: "Well tolerated"
    },

    FUROSEMIDE: {
    label: "Furosemide 12.5 mg",
    instructions: "Give your dog 1 tablet by mouth every 8 - 12 hours to drain fluid from the lungs.",
    class: "Diuretic",
    sideEffects: "May cause increased drinking and urination or bloodwork abnormalities"
    },

    FUROSEMIDEINJECTION: {
    label: "Furosemide injection",
    instructions: "Injection given in clinic to drain fluid from your dog’s lungs.",
    class: "Diuretic",
    sideEffects: "May cause increased drinking and urination or bloodwork abnormalities"
    },

    GABAPENTIN: {
    label: "Gabapentin 50mg",
    instructions: "Give your dog 1 tablet by mouth every 8 - 12 hours for treatment of pain.",
    class: "Analgesia, sedative",
    sideEffects: "May cause sedation"
    },

    GRAPIPRANT: {
    label: "Galliprant 100mg (grapiprant)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite. DO NOT USE WITHIN 3 DAYS OF OTHER NSAIDs OR STEROIDS."
    },

    HEARTGARD: {
    label: "Heartgard",
    instructions: "Give your dog 1 chewable tablet every 30 days for prevention of heartworms and common intestinal parasites.",
    class: "Parasiticide",
    sideEffects: "Rarely causes vomiting or diarrhea"
    },

    LIBRELAINJECTION: {
    label: "Librela injection (bedinvetmab)",
    instructions: "Medication injected beneath your dog’s skin to control arthritis over the next 4 weeks.",
    class: "Anti-arthritis (monoclonal antibody)",
    sideEffects: "Rarely causes seizures, urinary tract infection, or skin infections."
    },

    MELOXICAM: {
    label: "Meloxicam 7.5mg",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite. DO NOT USE WITH OTHER NSAIDs OR STEROIDS."
    },

    MELOXICAMLIQUID: {
    label: "Meloxicam liquid 1.5mg/mL",
    instructions: "Give your dog 1 mL by mouth every 24 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite. DO NOT USE WITHIN 3 DAYS OF OTHER NSAIDs OR STEROIDS."
    },

    NEXGARD: {
    label: "Nexgard",
    instructions: "Give your dog 1 chewable tablet every 30 days for prevention of fleas and ticks.",
    class: "Parasiticide",
    sideEffects: "Rarely causes vomiting or diarrhea"
    },

    PIMOBENDAN: {
    label: "Vetmedin 1.25mg\n(pimobendan)",
    instructions: "Give your dog 1 tablet by mouth every 12 hours to increase heart contractility & function.",
    class: "Inotropic agent",
    sideEffects: "Rarely causes vomiting (less than 1% of dogs)"
    },

    PROHEART6INJECTION: {
    label: "Proheart 6 (moxidectin)",
    instructions: "Medicine injected beneath your dog’s skin to prevent heartworm infection for 6 months.",
    class: "Antiparasitic",
    sideEffects: "May cause lethargy, decreased appetite, or vomiting"
    },

    PROHEART12INJECTION: {
    label: "Proheart 12 (moxidectin)",
    instructions: "Medicine injected beneath your dog’s skin to prevent heartworm infection for 12 months.",
    class: "Antiparasitic",
    sideEffects: "May cause lethargy, decreased appetite, or vomiting"
    },

    REVOLUTION: {
    label: "Revolution",
    instructions: "Apply contents between your dog’s ears every 30 days for prevention of heartworms, fleas, ticks, and common intestinal parasites.",
    class: "Parasiticide",
    sideEffects: "Rarely causes redness of the skin or hair loss"
    },

    SIMPARICATRIO: {
    label: "Simparica Trio",
    instructions: "Give your dog 1 chewable tablet by mouth every 30 days for prevention of heartworms, fleas, ticks, & common intestinal parasites.",
    class: "Antiparasitic",
    sideEffects: "Rarely causes vomiting, diarrhea, or neurologic abnormalities"
    },

    SYNOTICLIQUID: {
    label: "Synotic Otic Solution",
    instructions: "Starting today\nApply up to 5 drops in your dog’s affected ear every 12 hours for 1 week, then discontinue.",
    class: "Corticosteroid",
    sideEffects: "May cause short term ear discomfort or increased thirst/urination." },

    ZENRELIA: {
    label: "Zenrelia 15mg (ilunocitinib)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of allergies.",
    class: "Anti-allergy (JAK inhibitor)",
    sideEffects: "Well tolerated"
    }
    };

  // Medication Prefixes
    const MED_PREFIX = {
    START: "Starting today",
    CONTINUE: "Continue",
    WAIT3: "Wait 3 days then start",
    TOMORROW: "Starting tomorrow",
    DISCONTINUE: "Discontinue",
    CLINIC: "Administered in clinic ",
    NEWDOSE: "New dose",
    ASNEEDED: "As needed",
    };

    const PREFIX_ROW_COLOR = {
    CONTINUE: "#B6D7A8",     // green
    DISCONTINUE: "#EA9999",  // red
    ASNEEDED: "#FFE599",     // yellow
    NEWDOSE: "#B4A7D6",      // purple
    CLINIC: "#A4C2F4"        // blue
    // START, WAIT3, TOMORROW → default (null)
    };

    // Precompute every valid medication command
    const MED_COMMAND_LOOKUP = {};
    Object.keys(MEDICINE_REGISTRY).forEach(medKey => {
    Object.keys(MED_PREFIX).forEach(prefixKey => {
    const cmdKey = (medKey + prefixKey).toUpperCase();
    const med = MEDICINE_REGISTRY[medKey];
    const prefixText = MED_PREFIX[prefixKey];
    const color = PREFIX_ROW_COLOR[prefixKey] || null;

    MED_COMMAND_LOOKUP[cmdKey] = {
      rowData: [
        med.label,
        `${prefixText}\n${med.instructions}`,
        med.class,
        med.sideEffects
      ],
      color: color
    };
    });

    // Default START command if no prefix typed
    const defaultKey = (medKey + "START").toUpperCase();
    MED_COMMAND_LOOKUP[medKey.toUpperCase()] = MED_COMMAND_LOOKUP[defaultKey];
    });

    // Medication Command Processor
    function processMedicationCommand(keyword) {
    // Remove "/c" or "/C", uppercase, trim spaces
    const cmd = keyword.replace(/^\/c/i, "").toUpperCase().trim();
    return MED_COMMAND_LOOKUP[cmd] || null;
    }

/* ------------------ DIAGNOSIS & TEMPLATE BUFFER/RANKING ------------------ */
  // Diagnosis Registry & Rank
    const DIAGNOSIS_REGISTRY = {
    ATOPIC_DERMATITIS: {
    text: "Atopic dermatitis (allergies)",
    rank: 50
    },

    BLIND: {
    text: "Blind",
    rank: 80
    },

    CHERRY_EYE: {
    text: "Cherry eye",
    rank: 70
    },

    HEART_MURMUR: {
    text: "Heart murmur",
    rank: 1
    },

    OSTEOARTHRITIS: {
    text: "Osteoarthritis (arthritis)",
    rank: 20
    },

    OTITIS: {
    text: "Otitis externa (ear infection)",
    rank: 40
    },

    OVERWEIGHT: {
    text: "Overweight",
    rank: 70
    },

    PARTIALLY_BLIND: {
    text: "Partially blind",
    rank: 90
    },

    PARTIALLY_VACCINATED: {
    text: "Partially vaccinated",
    rank: 98
    },

    MILD_PERIODONTAL_DISEASE: {
    text: "Mild periodontal disease",
    rank: 30
    },

    MODERATE_PERIODONTAL_DISEASE: {
    text: "Moderate periodontal disease",
    rank: 30
    },

    PERIODONTAL_PERIODONTAL_DISEASE: {
    text: "Periodontal disease",
    rank: 30
    },

    SEVERE_PERIODONTAL_DISEASE: {
    text: "Severe periodontal disease",
    rank: 30
    },

    UNDERWEIGHT: {
    text: "Underweight",
    rank: 71
    },
    };

  // Shared: Insert With Format Break
    let diagnosisBuffer = [];
    let templateBuffer = [];

    function insertWithFormatBreak(body, index, insertFn) {
    const breaker = body.insertParagraph(index, " ");
    breaker.setBold(false);
    breaker.setUnderline(false);
    breaker.setItalic(false);
    breaker.setForegroundColor("#000000");
    breaker.setLineSpacing(1.0);

    const result = insertFn(index + 1);

    breaker.removeFromParent();
    return result;
    }

  // Add Diagnosis Codes
    function bufferDiagnoses(keys) {
    if (!keys) return;

    keys.forEach(key => {
    const diag = DIAGNOSIS_REGISTRY[key];
    if (!diag) return;

    if (!diagnosisBuffer.some(d => d.text === diag.text)) {
      diagnosisBuffer.push({
        text: diag.text,
        rank: diag.rank
      });
    }
    });
    }

  // Sort Diagnosis Buffer
    function sortDiagnosisBuffer() {
    diagnosisBuffer.sort((a, b) => a.rank - b.rank);
    }

    /* ------------------ Smart Insert Diagnoses ------------------ */
      function insertDiagnosesIntoDocument() {
      if (diagnosisBuffer.length === 0) return;

      const body = DocumentApp.getActiveDocument().getBody();

      // 1. Sort the incoming buffer by medical priority
      diagnosisBuffer.sort((a, b) => (a.rank || 99) - (b.rank || 99));

      // 2. Find the "Diagnosis" header index
      let headerIndex = -1;
      let formatting = {};

      for (let i = 0; i < body.getNumChildren(); i++) {
      const element = body.getChild(i);
      if (element.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

      const paragraph = element.asParagraph();
      const text = paragraph.getText().trim();

      if (text.toLowerCase() !== "diagnosis") continue;

      const textObj = paragraph.editAsText();
      const color = textObj.getForegroundColor(0);
      const isGreen = (color === "#008000" || color === "#6aa84f" || color === "#b6d7a8");

      if (textObj.isBold(0) && textObj.isUnderline(0) && !isGreen) {
      headerIndex = i;
      // Capture formatting to apply to new rows
      formatting = {
      indentFirstLine: paragraph.getIndentFirstLine(),
      indentStart: paragraph.getIndentStart(),
      lineSpacing: paragraph.getLineSpacing()
      };
      break;
      }
      }

      // If we can't find the Diagnosis header, we can't safely insert
      if (headerIndex === -1) return;

      // 3. Process each diagnosis in the buffer
      diagnosisBuffer.forEach(newDiag => {
      let inserted = false;
      const newRank = newDiag.rank || 99;

      // Search paragraphs following the header
      for (let j = headerIndex + 1; j < body.getNumChildren(); j++) {
      const child = body.getChild(j);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) break; // Stop if we hit a table/other

      const pText = child.asParagraph().getText().trim();

      // If we hit an empty line or the next section, stop and insert here
      if (!pText) {
      const newPara = body.insertParagraph(j, newDiag.text);
      applyDiagFormatting(newPara, formatting);
      inserted = true;
      break;
      }

      // Identify the rank of the diagnosis already sitting on this line
      let existingRank = null;
      for (const key in DIAGNOSIS_REGISTRY) {
      if (DIAGNOSIS_REGISTRY[key].text === pText) {
      existingRank = DIAGNOSIS_REGISTRY[key].rank || 99;
      break;
      }
      }

      // If we found a match and it's lower priority (higher rank #), cut in line!
      if (existingRank !== null && existingRank > newRank) {
      const newPara = body.insertParagraph(j, newDiag.text);
      applyDiagFormatting(newPara, formatting);
      inserted = true;
      break;
      }

      // If the text exists but isn't in our registry, it's likely a custom note. 
      // We continue scanning until we find a clear "Rank" match or the end of the list.
      }

      // 4. Fallback: If we didn't find a spot to cut in, append to the end of the section
      if (!inserted) {
      // Find the end of the diagnosis block (first empty line or end of doc)
      let searchIdx = headerIndex + 1;
      while (searchIdx < body.getNumChildren() && body.getChild(searchIdx).asParagraph().getText().trim() !== "") {
      searchIdx++;
      }
      const newPara = body.insertParagraph(searchIdx, newDiag.text);
      applyDiagFormatting(newPara, formatting);
      }
      });

      diagnosisBuffer = [];
      }

      /**
       * Helper to apply the captured header formatting to new diagnosis rows
       */
      function applyDiagFormatting(paragraph, fmt) {
      paragraph.setIndentFirstLine(fmt.indentFirstLine);
      paragraph.setIndentStart(fmt.indentStart);
      paragraph.setLineSpacing(fmt.lineSpacing);
      }

  // Template Buffer
    function bufferTemplate(templateObj, rank = 99) {
    if (!templateObj || !templateObj.text) return;

    if (!templateBuffer.some(t => t.text === templateObj.text)) {
    templateBuffer.push({ ...templateObj, rank: rank });
    }
    }

    /* ------------------ Deep Fuzzy Diagnosis-Linked Insert ------------------ */
      function insertTemplatesIntoDocument() {
      if (templateBuffer.length === 0) return;

      const body = DocumentApp.getActiveDocument().getBody();
      const allDiags = DIAGNOSIS_REGISTRY;

      // 1. Sort incoming buffer by rank
      templateBuffer.sort((a, b) => (a.rank || 999) - (b.rank || 999));

      // 2. Find the "Comprehensive Summary" anchor
      let summaryIndex = -1;
      for (let i = 0; i < body.getNumChildren(); i++) {
      const child = body.getChild(i);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

      const text = child.asParagraph().getText().trim().toLowerCase();
      if (text === "comprehensive summary") {
      summaryIndex = i;
      break;
      }
      }

      // --- CLEANUP WHITESPACE ---
      if (summaryIndex !== -1) {
      let nextIdx = summaryIndex + 1;
      while (nextIdx < body.getNumChildren()) {
      const nextChild = body.getChild(nextIdx);
      if (nextChild.getType() === DocumentApp.ElementType.PARAGRAPH && 
      nextChild.asParagraph().getText().trim() === "") {
      if (body.getNumChildren() > 1) {
      try {
      nextChild.removeFromParent();
      } catch (e) { break; }
      } else { break; }
      } else { break; }
      }
      }

      // Define where the summary section starts
      const targetBase = (summaryIndex === -1) ? body.getNumChildren() : summaryIndex + 1;

      templateBuffer.forEach(newTmpl => {
      let inserted = false;
      const newRank = newTmpl.rank || 999;

      // Search existing paragraphs to see where to insert based on rank
      for (let i = targetBase; i < body.getNumChildren(); i++) {
      const child = body.getChild(i);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

      const p = child.asParagraph();
      const pText = p.getText().trim().toLowerCase();

      // If we hit an empty line, the summary section has ended
      if (!pText) break; 

      if (!p.editAsText().isBold(0) || !pText.includes(':')) continue;

      const headerText = pText.split(':')[0]; 
      let existingRank = null;

      for (const key in allDiags) {
      const diagText = allDiags[key].text.toLowerCase();
      const diagWords = diagText.split(/\s+/).filter(word => word.length > 2);
      const matchCount = diagWords.filter(word => headerText.includes(word)).length;
      if (matchCount >= Math.ceil(diagWords.length * 0.7)) {
      existingRank = allDiags[key].rank;
      break;
      }
      }

      if (existingRank !== null && existingRank > newRank) {
      insertTemplateAtIndex(body, newTmpl, i);
      inserted = true;
      break;
      }
      }

      // --- FIXED FALLBACK LOGIC ---
      if (!inserted) {
      // Instead of going to the absolute end of the doc, find the end of the text block
      let fallbackIdx = targetBase;
      while (fallbackIdx < body.getNumChildren()) {
      const next = body.getChild(fallbackIdx);
      // Stop if we hit a non-paragraph or an empty paragraph (the "floor")
      if (next.getType() !== DocumentApp.ElementType.PARAGRAPH || next.asParagraph().getText().trim() === "") {
      break;
      }
      fallbackIdx++;
      }
      insertTemplateAtIndex(body, newTmpl, fallbackIdx);
      }
      });

      templateBuffer = [];
      }

/* ------------------ FORMAT REGISTRY ------------------ */
  // Reset Registry
    const FORMAT_REGISTRY = {  
    DOG_VETERINARY_VISIT_SUMMARY:
    `🐶 Your Dog's Veterinary Visit Summary 🐕`,

    CAT_VETERINARY_VISIT_SUMMARY:
    `😺 Your Cat's Veterinary Visit Summary 🐈`,

    VETERINARIAN:
    `Veterinarian`,
  
    DIAGNOSTICS:
    `Diagnostics`,

    PHYSICAL_EXAM:
    `Physical exam`,

    HW_TEST:
    `Heartworm & tickborne disease test:`,

    FeLV_FIV_TEST:
    `Feline leukemia virus & FIV test:`,

    PARVOVIRUS_TEST:
    `Parvovirus test:`,

    EAR_CYTOLOGY:
    `Ear cytology:`,

    IMPRESSION_SMEAR:
    `Impression smear:`,

    FINE_NEEDLE_ASPIRATE:
    `Fine needle aspirate:`,
  
    SCHIRMER_TEAR_TEST:
    `Schirmer tear test:`,

    FLUORESCEIN_EYE_STAIN:
    `Fluorescein eye stain:`,

    TONOMETRY:
    `Tonometry:`,
  
    CBC_TEST:
    `Complete blood count:`,

    CHEMISTRIES:
    `Plasma biochemical analysis:`,

    TOTAL_T4:
    `Total T4 (Thyroid hormone test):`,

    URINALYSIS:
    `Urinalysis:`,

    FECAL_TEST:
    `Fecal test:`,

    PROBNP_TEST:
    `ProBNP:`,
  
    THORACIC_RADIOGRAPHS:
    `Thoracic radiographs:`,

    ABDOMINAL_RADIOGRAPHS:
    `Abdominal radiographs:`,

    EXTREMITY_RADIOGRAPHS:
    `Extremity radiographs:`,

    ABDOMINAL_ULTRASOUND:
    `Abdominal ultrasound:`,

    ECHOCARDIOGRAM:
    `Echocardiogram:`,
  
    DIAGNOSIS_RESET:
    `Diagnosis`,
  
    SUMMARY_RESET:
    `Comprehensive Summary`,

    RESULTS_PENDING:
    `Results pending`,

  // Vaccines
    BORDETELLA_VXN:
    'The 1 year bordetella vaccine',

    DAPP_INITIAL:
    'The initial distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine',

    DAPP_BOOSTER:
    'The distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine',

    DAPP_1YR:
    'The 1 year distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine',

    DAPP_3YR:
    'The 3 year distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine',

    IMMEDIATELY: g =>
    `bring your ${g.dog} back immediately for treatment`,

    LEPTO_INITIAL:
    'the initial lepto vaccine',

    LEPTO_VXN:
    /the 1 year lepto vaccine/i,

    NEXT_APPOINTMENT_HEADER: 'Next appointment:',

    PUPPY_QUARANTINE: g =>
    `Because your ${g.dog} ${g.is} still getting vaccines to better ${g.his} immune system, it’s best to keep ${g.him} away from dog parks & other dogs that aren’t part of your household until two weeks after ${g.he} ${g.has} finished ${g.his} puppy vaccines.`,

    Quarantine_12WK: g =>
    `Keep ${g.him} away from dog parks, training facilities, and other dogs until then.`,

    Quarantine_16WK: g =>
    `During this time, keep ${g.him} away from dog parks & other dogs that aren’t part of your household.`,

    RABIES_1YR:
    'The 1 year rabies vaccine',

    RABIES_3YR:
    'The 3 year rabies vaccine',

    RARE_RXN: g =>
    `These reactions are rare & not expected to occur in your ${g.dog}.`,

    VACCINES_HEADER:
    'Vaccines:',

    VXN_BOOSTER: g =>
    `Your ${g.dog} will need a booster of the DAPP and lepto vaccines in 3 - 4 weeks.`,

    VXN_RXN: g =>
    `Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads).`,

  // Labwork
    HEARTWORMS_HEADER: 
    'Heartworms:',

    HEARTWORM_TEST: g => 
    `A heartworm test was performed on your ${g.dog}. We will contact you in 3 - 4 business days with the results.`,

    LAB_RESULTS: g => 
    `Samples were drawn from your ${g.dog}. You will receive a call in 3 - 4 business days with the results.`,
    
    LABWORK: 
    'Early detection labwork:',

    HW_PREVENTION_SENTENCE: g =>
    `Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your ${g.dog} on monthly preventatives such as Heartgard, Simparica Trio, Revolution, etc.`,

  // Spay & Neuter
    LARGE_NEUTER: g =>
    `It is recommended you have ${g.him} neutered once ${g.he} ${g.is} 10 - 12 months old if you don’t intend to breed ${g.him}.`,
    
    NEUTER_HEADER:
    'Neuter:',

    SMALL_NEUTER: g =>
    `It is recommended you have ${g.him} neutered once ${g.he} ${g.is} 6 months old if you don’t intend to breed ${g.him}.`,
    
    SPAY1: g =>
    `It is recommended you have your ${g.dog} spayed at 6 months if you do not intend to breed ${g.him}.`,

    SPAY2:
    '1 in 4 female dogs will get breast cancer if they are not spayed by their second heat cycle. In dogs there is a 50% chance breast cancer spreads throughout the body.',

    SPAY3: g =>
    `Even if ${g.he} ${g.has} already had ${g.his} second heat cycle, spaying is still recommended since many mammary tumors are stimulated by estrogen & pyometra is still a possibility.`,

    SPAY_HEADER: 'Spay:',

  // Diet
    DIET_HEADER:
    'Food:',

    GRAIN_FREE:
    'It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites.',

    HILLS_DOG_DRY_LINK: {
    text: "Hill's dog dry food",
    url: 'https://www.hillspet.com/dog-food?lifestage=adult&productform=dry'
    },

    HILLS_DOG_WET_LINK: {
    text: "Hill's dog wet food",
    url: 'https://www.hillspet.com/dog-food?lifestage=adult&productform=canned&productform=stew'
    },

    HILLS_PUPPY_DRY_LINK: {
    text: "Hill's puppy dry food",
    url: 'https://www.hillspet.com/dog-food?lifestage=puppy&productform=dry'
    },

    HILLS_PUPPY_WET_LINK: {
    text: "Hill's puppy wet food",
    url: 'https://www.hillspet.com/dog-food?lifestage=puppy&productform=canned&productform=stew'
    },

    HILLS_SR_DOG_DRY_LINK: {
    text: "Hill's senior dog dry food",
    url: 'https://www.hillspet.com/dog-food?lifestage=mature&lifestage=senior&productform=dry'
    },

    HILLS_SR_DOG_WET_LINK: {
    text: "Hill's senior dog wet food",
    url: 'https://www.hillspet.com/dog-food?lifestage=mature&lifestage=senior&productform=canned&productform=stew'
    },

    PURINA_DOG_DRY_LINK: {
    text: 'Purina dog dry food',
    url: 'https://www.purina.com/dogs/dog-food/dry?items_per_page=10&sort_by=relevance&f%5B0%5D=life-stage%3A1504'
    },

    PURINA_DOG_WET_LINK: {
    text: 'Purina dog wet food',
    url: 'https://www.purina.com/dogs/dog-food/wet?items_per_page=10&sort_by=relevance&f%5B0%5D=life-stage%3A1504'
    },

    PURINA_PUPPY_DRY_LINK: {
    text: 'Purina puppy dry food',
    url: 'https://www.purina.com/dogs/dog-food/dry/puppy-food'
    },

    PURINA_PUPPY_WET_LINK: {
    text: 'Purina puppy wet food',
    url: 'https://www.purina.com/dogs/dog-food/wet/puppy-food'
    },

    PURINA_SR_DOG_DRY_LINK: {
    text: 'Purina senior dog dry food',
    url: 'https://www.purina.com/dogs/dog-food/senior?f%5B0%5D=category%3A14&f%5B1%5D=life-stage%3A1503&items_per_page=10&sort_by=relevance'
    },

    PURINA_SR_DOG_WET_LINK: {
    text: 'Purina senior dog wet food',
    url: 'https://www.purina.com/dogs/dog-food/senior?f%5B0%5D=category%3A13&f%5B1%5D=life-stage%3A1503&items_per_page=10&sort_by=relevance'
    },

    ROYAL_CANIN_DOG_DRY_LINK: {
    text: 'RC dog dry food',
    url: 'https://www.royalcanin.com/us/dogs/products/adult-dog-food?lifestage=adult&digital_sub_category=dry_food'
    },

    ROYAL_CANIN_DOG_WET_LINK: {
    text: 'RC dog wet food',
    url: 'https://www.royalcanin.com/us/dogs/products/adult-dog-food?lifestage=adult&digital_sub_category=wet_food'
    },

    ROYAL_CANIN_PUPPY_DRY_LINK: {
    text: 'RC puppy dry food',
    url: 'https://www.royalcanin.com/us/dogs/products/puppy-food?lifestage=baby|junior|puppy&digital_sub_category=dry_food'
    },

    ROYAL_CANIN_PUPPY_WET_LINK: {
    text: 'RC puppy wet food',
    url: 'https://www.royalcanin.com/us/dogs/products/puppy-food?lifestage=baby|junior|puppy&digital_sub_category=wet_food'
    },

    ROYAL_CANIN_SR_DOG_DRY_LINK: {
    text: 'RC senior dog dry food',
    url: 'https://www.royalcanin.com/us/dogs/products/senior-dog-food?lifestage=ageing|mature&digital_sub_category=dry_food',
    },

    ROYAL_CANIN_SR_DOG_WET_LINK: {
    text: 'RC senior dog wet food',
    url: 'https://www.royalcanin.com/us/dogs/products/senior-dog-food?lifestage=ageing|mature&digital_sub_category=wet_food'
    },

  // Dental
    BRUSHING_LESS_EFFECTIVE: 
    'Brushing can still be performed right now but will be most effective after the next cleaning.',

    COHAT_RECOMMENDED: g => 
    `Your ${g.dog} needs a Complete Oral Health Assessment and Treatment (COHAT) procedure.`,

    DENTAL_BRUSHING: 
    'Brushing the outside for 1.5 seconds is more than enough.',

    DENTAL_BRUSHING_CORE: 
    /The best way to keep your .*? teeth healthy is to brush them daily for 10 seconds total using a .*?\./,
    
    DENTAL_HEADER: 'Dental care:',

    LARGE_TOOTHBRUSH_LINK: {
    text: 'medium/large dog toothbrush',
    url: 'https://a.co/d/0dgjdsJ8'
    },

    MILD_DENTAL_DZ: g =>
    `Your ${g.dog} ${g.has} early dental disease.`,

    PERIODONTAL_DISEASE_HEADER: 
    'Periodontal disease:',

    SCHEDULE_DENTAL: 
    'Schedule a dental cleaning within the next three months.',

    SMALL_DOG_TOOTHBRUSH_LINK: {
    text: 'small dog toothbrush',
    url: 'https://a.co/d/00v6EBUQ'
    },

    SOFTEN_FOOD: 
    `In the meantime, feed soft food such as wet food or dry food soaked in a few tablespoons of warm water 30 seconds prior to feeding.`,

    TOOTHPASTE_LINK: {
    text: 'animal safe toothpaste',
    url: 'https://a.co/d/07LxFqth'
    },

    VOHC_DOG_LINK: {
    text: 'Veterinary Oral Health Council website',
    url: 'https://vohc.org/accepted-products/#dogs'
    },

    XYLITOL: 
    '(make sure xylitol isn’t listed as an ingredient),',

  // Weight Management
    DIET_LIFESPAN:
    /Helping .*? to lose weight can increase .*? life span by as much as 1 ½ years\./,

    DIET2_LIFESPAN:
    /Continuing to help .*? lose weight can extend .*? life span by as much as 1 ½ years\./,
    
    DIET_WEEKLY_GOAL:
    /We’re aiming to have .*? lose 1 - 2% of .*? body weight per week\./,

    HEALTHY_WEIGHT_LIFESPAN: g =>
    `Your ${g.dog} is a healthy weight for a ${g.dog} of ${g.his} size.`,

    HILLS_DOG_DIET_DRY_LINK: {
    text: "Hill’s weight loss dry food",
    url: 'https://www.hillspet.com/dog-food?productform=dry&condition=weightmanagement'
    },

    HILLS_DOG_DIET_WET_LINK: {
    text: "Hill’s weight loss wet food",
    url: 'https://www.hillspet.com/dog-food?&productform=canned&productform=stew&condition=weightmanagement'
    },

    OVERWEIGHT_WARNING:
    /Your .*? weighs? more than the average .*? of .*? size\./,

    OVERWEIGHT_DOG_SIGNS: g =>
    `Signs of an overweight ${g.dog}`,

    PURINA_DOG_DIET_DRY_LINK: {
    text: 'Purina weight loss dry food',
    url: 'https://www.purina.com/dogs/dog-food/dry?items_per_page=10&sort_by=relevance&f%5B0%5D=health_benefits%3A940'
    },

    PURINA_DOG_DIET_WET_LINK: {
    text: 'Purina weight loss wet food',
    url: 'https://www.purina.com/dogs/dog-food/wet?items_per_page=10&sort_by=relevance&f%5B0%5D=health_benefits%3A940'
    },

    ROYAL_CANIN_DOG_DIET_DRY_LINK: {
    text: 'RC weight loss dry',
    url: 'https://www.royalcanin.com/us/dogs/products/retail-products?digital_sub_category=dry_food&specific_needs=weight_management'
    },

    ROYAL_CANIN_DOG_DIET_WET_LINK: {
    text: 'RC weight loss wet',
    url: 'https://www.royalcanin.com/us/dogs/products/retail-products?specific_needs=weight_management&digital_sub_category=wet_food'
    },
    
    UNDERWEIGHT_DOG_SIGNS: g =>
    `Signs of an underweight ${g.dog}`,

    UNDERWEIGHT_HEADER:
    `Underweight:`,

    UNDERWEIGHT_LIFESPAN: g =>
    `Helping ${g.him} gain weight can increase ${g.his} quality of life.`,

    UNDERWEIGHT_GOAL: g =>
    `We're aiming to have ${g.him} gain approximately 10% of ${g.his} current weight. Failure to gain weight is concerning for disease and would prompt us to perform tests such as labwork, ultrasound, or x-rays.`,

    UNDERWEIGHT_WARNING: g =>
    `Your ${g.dog} weighs less than the average ${g.dog} of ${g.his} size.`,

    WEIGHT_HEADER:
    'Weight:',

  // Ophthalmology
    BLIND_HEADER:
    "Blind:",

    BLIND_OBSTACLE_COURSE: g =>
    `Make sure to keep your living space free of obstacles to prevent your ${g.dog} from tripping or bumping into things by mistake.`,

    CHERRY_EYE_HEADER: g => 
    `Cherry ${g.eye}:`,

    CHERRY_EYE_SURGERY_RECOMMENDATION: g => 
    `For dogs older than 1 year it is best to have the cherry ${g.eye} corrected as soon as possible.`,

    HALO_HARNESS_ARTICLE: {
      text: "Halo harness",
      url: `https://www.muffinshalo.com/`
    },
  
  // Cardiology
    HEART_MURMUR_ARTICLE: {
    text: "Heart Murmurs in Dogs and Cats article",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4952593`
    },

    HEART_MURMUR_HEADER:
    "Heart murmur:",

    HEART_MURMUR_HEARD: g =>
    `A heart murmur was heard in your ${g.dog} today.`,

    HEART_MURMUR_GRADING:
    "A higher grade (5 & 6) does not always indicate worse disease & a lower grade (1 & 2) does not always indicate a better disease.",

    HEART_MURMUR_WARNING_SIGNS: g =>
    `If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease.`,

    SCHEDULE_ECHO: g =>
    `an echocardiogram to look at the inner workings of the heart and diagnose the cause of the disease will need to be scheduled.`,

    SCHEDULE_ECHOS: g =>
    `An echocardiogram to look at the inner workings of the heart and diagnose the cause of the disease will need to be scheduled.`,

  // Musculoskeletal
    ARTHRITIS_DETECTED: g =>
    `Arthritis was detected in your ${g.dogs} joints.`,

    ARTHRITIS_WEIGHT_MANAGEMENT: g =>
    `Keeping your ${g.dog} an appropriate weight can also help reduce joint pain.`,

    GABAPENTIN_ARTHRITIS_MANAGEMENT: g =>
    `Your ${g.dog} ${g.is} known to have arthritis & ${g.is} currently on gabapentin.`,

    GABAPENTIN_DECISION: 
    `You’ve elected to try gabapentin`,

    JOINT_SUPPLEMENTS_DECISION: 
    `At this time you’ve elected to try joint supplements.`,
    
    JOINT_SUPPLEMENTS_KNOWN: g =>
    `Your ${g.dog} ${g.is} known to have arthritis & currently gets joint supplements.`,

    LIBRELA_ADVERSE_RXN: 
    `Watch for signs of reaction including vomiting, diarrhea, lethargy, & excessive panting/fever.`,

    NSAID_LABWORK_REQUIREMENT: 
    `You’ve elected to try a non-steroidal anti-inflammatory drug (NSAID) to reduce pain & inflammation from arthritis. Bloodwork is required every 6 - 12 months while on NSAIDs.`,

    NSAID_MONITORING: 
    `Bloodwork is required every 6 - 12 months while on NSAIDs.`,

    OSTEOARTHRITIS_HEADER: 
    `Osteoarthritis:`,

  // Dermatology
    ALLERGY_EAR_RELATIONSHIP:
    `In fact, one of the most common causes of chronic ear infections is allergies.`,
    
    APOQUEL_STARTER:
    `Apoquel has been sent home to resolve allergies. Give as prescribed.`,

    ATOPIC_DERMATITIS_HEADER:
    `Atopic Dermatitis:`,

    ANTIHISTAMINE_DOSAGE1:
    `You can give over the counter antihistamines such as Benadryl 25mg (give up to 1 tablet per 25 lbs every 12 hours) or Zyrtec 10mg (give up to 1 tablet per 10 lbs every 12 - 24 hours).`,

    ANTIHISTAMINE_DOSAGE2:
    `Continue to give Benadryl 25mg (give up to 1 tablet per 25 lbs every 12 hours) or Zyrtec 10mg (give up to 1 tablet per 10 lbs every 12 - 24 hours) as needed.`,

    ANTIHISTAMINE_ADDITION:
    'You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.',

    CYTOPOINT_ADDITIONAL_SUPPORT: g =>
    `If the allergies return before 4 weeks, your ${g.dog} may need Apoquel or Zenrelia in addition to Cytopoint.`,
    
    CYTOPOINT_STARTER:
    `Cytopoint has been given in clinic to resolve allergies.`,

    SYNERGISTIC_MEDICINE:
    `These medications improve the effectiveness of the other and are safe to give together.`,

    ZENRELIA_AND_APOQUEL_WARNING:
    `Give daily for 1 month for best results. Do not give Zenrelia in the same 24 hours as Apoquel.`,

    ZENRELIA_STARTER:
    `Zenrelia has been sent home to resolve allergies. Give as prescribed.`,

  // Immunology
    ADENOVIRUS_LINK: {
    text: `adenovirus & parainfluenza article`,
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951478`,
    },

    BORDETELLA_HEADER:
    `Bordetella:`,

    BORDETELLA_LINK: {
    text: `Kennel Cough in Dogs`,
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951478`,
    },

    BORDETELLA_NAME:
    `Bordetella,`,

    DAPP_HEADER:
    `DAPP:`,

    DAPP_WARNING:
    `Some studies suggest that distemper can be spread to humans. There is no cure for distemper.`,

    DISTEMPER_LINK: {
    text: `distemper article`,
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=11692001`,
    },

    LEPTO_HEADER:
    `Lepto:`,

    LEPTO_LINK: {
    text: `Leptospirosis in Dogs`,
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951453`,
    },

    LEPTO_WARNING: g =>
    `Lepto can quickly kill our ${g.dogs} & can be spread to humans.`,

    PARVO_LINK: {
    text: `parvovirus article`,
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951460`,
    },

    RABIES_CURE: g =>
    `THERE IS NO CURE FOR RABIES. The only way to test for rabies involves decapitating an animal & taking samples of the brain. State law requires any ${g.dog} that is exposed to rabies to either undergo quarantine for up to 6 months or be euthanized.`,

    RABIES_HEADER:
    `Rabies:`,

    RABIES_IN_ANIMALS_LINK: {
    text: `Rabies in Animals article`,
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951479`,
    },

    TEXAS_RABIES_LINK: {
    text: `the Texas government website`,
    url: `https://www.dshs.texas.gov/notifiable-conditions/zoonosis-control/zoonosis-control-diseases-and-conditions/rabies`,
    },

  // General Illness
    COMMON_CAUSES:
    /Common causes/i,

    DIAGNOSIS:
    /Diagnosis/i,
    
    DIAGNOSE:
    /Diagnose/i,

    SYMPTOMS:
    /Symptoms/i,

    TREATMENT:
    /Treatment/i,

    };

/* ------------------ RESET & GENERIC ------------------ */
  // Canine Reset
    function generateCanineResetTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `🐶 Your Dog's Veterinary Visit Summary 🐕`,
    `Veterinarian`,
    `Dr. Osadiaye (Oh-sah-dya-yay)`,
    ``,
    `Diagnostics`,
    `Physical exam`,
    `Heartworm & tickborne disease test: Results pending`,
    `Feline leukemia virus & FIV test: Results pending`,
    `Parvovirus test: Results pending`,
    `Ear cytology: Results pending`,
    `Impression smear: Results pending`,
    `Fine needle aspirate: Results pending`,
    ``,
    `Schirmer tear test: Results pending`,
    `Fluorescein eye stain: Results pending`,
    `Tonometry: Results pending`,
    ``,
    `Complete blood count: Results pending`,
    `Plasma biochemical analysis: Results pending`,
    `Total T4 (Thyroid hormone test): Results pending`,
    `Urinalysis: Results pending`,
    `Fecal test: Results pending`,
    `ProBNP: Results pending`,
    ``,
    `Thoracic radiographs: Results pending`,
    `Abdominal radiographs: Results pending`,
    `Extremity radiographs: Results pending`,
    `Abdominal ultrasound: Results pending`,
    `Echocardiogram: Results pending`,
    ``,
    `Diagnosis`,
    ``,
    ``,
    `Comprehensive Summary`,
    ].join('\n');

    return {
    sex,
    text,
    blockLineSpacing: 1.0,  // Entire template single spaced
    titleKeys: [
    `DOG_VETERINARY_VISIT_SUMMARY`,
    ],

    boldKeys: [
    `PHYSICAL_EXAM`,
    `HW_TEST`,
    `FeLV_FIV_TEST`,
    `PARVOVIRUS_TEST`,
    `EAR_CYTOLOGY`,
    `IMPRESSION_SMEAR`,
    `FINE_NEEDLE_ASPIRATE`,
    `SCHIRMER_TEAR_TEST`,
    `FLUORESCEIN_EYE_STAIN`,
    `TONOMETRY`,
    `CBC_TEST`,
    `CHEMISTRIES`,
    `TOTAL_T4`,
    `URINALYSIS`,
    `FECAL_TEST`,
    `PROBNP_TEST`,
    `THORACIC_RADIOGRAPHS`,
    `ABDOMINAL_RADIOGRAPHS`,
    `EXTREMITY_RADIOGRAPHS`,
    `ABDOMINAL_ULTRASOUND`,
    `ECHOCARDIOGRAM`,
    ],

    boldUnderlineKeys: [
    `VETERINARIAN`,
    `DIAGNOSTICS`,
    `DIAGNOSIS_RESET`,
    `SUMMARY_RESET`,
    ],

    italicKeys: [
    'RESULTS_PENDING',
    ],

    doubleSpacedKeys: [
    'SUMMARY_RESET',
    ],
    };
    }

  // Feline Reset
    function generateFelineResetTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `😺 Your Cat's Veterinary Visit Summary 🐈`,
    `Veterinarian`,
    `Dr. Osadiaye (Oh-sah-dya-yay)`,
    ``,
    `Diagnostics`,
    `Physical exam`,
    `Heartworm & tickborne disease test: Results pending`,
    `Feline leukemia virus & FIV test: Results pending`,
    `Parvovirus test: Results pending`,
    `Ear cytology: Results pending`,
    `Impression smear: Results pending`,
    `Fine needle aspirate: Results pending`,
    ``,
    `Schirmer tear test: Results pending`,
    `Fluorescein eye stain: Results pending`,
    `Tonometry: Results pending`,
    ``,
    `Complete blood count: Results pending`,
    `Plasma biochemical analysis: Results pending`,
    `Total T4 (Thyroid hormone test): Results pending`,
    `Urinalysis: Results pending`,
    `Fecal test: Results pending`,
    `ProBNP: Results pending`,
    ``,
    `Thoracic radiographs: Results pending`,
    `Abdominal radiographs: Results pending`,
    `Extremity radiographs: Results pending`,
    `Abdominal ultrasound: Results pending`,
    `Echocardiogram: Results pending`,
    ``,
    `Diagnosis`,
    ``,
    ``,
    `Comprehensive Summary`,
    ].join('\n');

    return {
    sex,
    text,
    blockLineSpacing: 1.0,  // Entire template single spaced
    titleKeys: [
    `CAT_VETERINARY_VISIT_SUMMARY`,
    ],

    boldKeys: [
    `PHYSICAL_EXAM`,
    `HW_TEST`,
    `FeLV_FIV_TEST`,
    `PARVOVIRUS_TEST`,
    `EAR_CYTOLOGY`,
    `IMPRESSION_SMEAR`,
    `FINE_NEEDLE_ASPIRATE`,
    `SCHIRMER_TEAR_TEST`,
    `FLUORESCEIN_EYE_STAIN`,
    `TONOMETRY`,
    `CBC_TEST`,
    `CHEMISTRIES`,
    `TOTAL_T4`,
    `URINALYSIS`,
    `FECAL_TEST`,
    `PROBNP_TEST`,
    `THORACIC_RADIOGRAPHS`,
    `ABDOMINAL_RADIOGRAPHS`,
    `EXTREMITY_RADIOGRAPHS`,
    `ABDOMINAL_ULTRASOUND`,
    `ECHOCARDIOGRAM`,
    ],

    boldUnderlineKeys: [
    `VETERINARIAN`,
    `DIAGNOSTICS`,
    `DIAGNOSIS_RESET`,
    `SUMMARY_RESET`,
    ],

    italicKeys: [
    'RESULTS_PENDING',
    ],

    doubleSpacedKeys: [
    'SUMMARY_RESET',
    ],
    };
    }

/* ------------------ Reverse Template Generator ------------------ */
  // Main Function
    function reverseGenerateTemplate() {
    const body = DocumentApp.getActiveDocument().getBody();
    const reverseMap = getReverseRegistryMap(); // Maps existing text/urls to keys

    if (typeof MEDICINE_REGISTRY === 'undefined') { var MEDICINE_REGISTRY = {}; }

    const NEW_LINKS_REGISTRY = {};
    const boldKeys = new Set();
    const boldUnderlineKeys = new Set();
    const italicKeys = new Set();
    const greenKeys = new Set();
    const redKeys = new Set();
    const linkKeys = new Set();
    const paragraphs = [];

    for (let i = 0; i < body.getNumChildren(); i++) {
    const element = body.getChild(i);
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
    const paragraph = element.asParagraph();
    let text = paragraph.getText() || "";
    if (!text.trim() || text.includes("/* --- REVERSE")) continue;

    paragraphs.push("`" + escapeBackticks(convertToGPronouns(text)) + "`");

    extractFormattingSpans(paragraph.editAsText(), {
    boldKeys, boldUnderlineKeys, italicKeys, greenKeys, redKeys, linkKeys
    }, reverseMap, NEW_LINKS_REGISTRY);
    }

    if (element.getType() === DocumentApp.ElementType.TABLE) {
    const table = element.asTable();
    for (let r = 1; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    const rowData = [];
    for (let c = 0; c < row.getNumCells(); c++) {
    rowData.push((row.getCell(c).getText() || "").trim());
    }
    const label = rowData[0] || "";
    let drugBase = label.split(' ')[0].replace(/[^A-Za-z]/g, "").toUpperCase();
    if (!drugBase) continue;

    let drugName = drugBase;
    if (label.toLowerCase().includes("injection")) drugName += "INJECTION";
    else if (label.toLowerCase().includes("liquid") || label.toLowerCase().includes("oral susp")) drugName += "LIQUID";

    if (!MEDICINE_REGISTRY[drugName]) {
    MEDICINE_REGISTRY[drugName] = {
    label: label, 
    instructions: rowData[1] || "",
    class: rowData[2] || "Unknown", 
    sideEffects: rowData[3] || "Unknown"
    };
    }
    }
    }
    }

    const formatKeys = (set) => Array.from(set).sort().map(k => `"${k}"`).join(",\n      ");

  // Generate Registry Code (New links only)
    let registryCode = "";
    const newLinkKeys = Object.keys(NEW_LINKS_REGISTRY);
    if (newLinkKeys.length > 0) {
    registryCode += "\n/* ------------------ New Format Registry Entries ------------------ */\n";
    newLinkKeys.sort().forEach(key => {
    const link = NEW_LINKS_REGISTRY[key];
    registryCode += `${key} : {\n  text: "${String(link.text).replace(/"/g, '\\"')}",\n  url: \`${link.url}\`\n},\n`;
    });
    }

  // Prescription Registry
    let prescripCode = "\n/* ------------------ Updated Prescription Registry ------------------ */\nconst PRESCRIPTION_REGISTRY = {\n";
    Object.keys(MEDICINE_REGISTRY).sort().forEach(key => {
    const med = MEDICINE_REGISTRY[key];
    const safeLabel = String(med.label || "").replace(/"/g, '\\"');
    const safeInstr = String(med.instructions || "").replace(/"/g, '\\"').replace(/\n/g, "\\n");
    prescripCode += `  ${key}: { label: "${safeLabel}", instructions: "${safeInstr}", class: "${med.class}", sideEffects: "${med.sideEffects}" },\n`;
    });
    prescripCode += "};\n";

  // Generated Template
    let templateCode = `// Generated Template
    function generateTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    text: [
    ${paragraphs.join(",\n      ")}
    ].join('\\n'),
    diagnoses: [""],\n`;

    // Only add keys if they have content
    if (boldKeys.size > 0) templateCode += `    boldKeys: [\n      ${formatKeys(boldKeys)}\n    ],\n\n`;
    if (boldUnderlineKeys.size > 0) templateCode += `    boldUnderlineKeys: [\n      ${formatKeys(boldUnderlineKeys)}\n    ],\n\n`;
    if (italicKeys.size > 0) templateCode += `    italicKeys: [\n      ${formatKeys(italicKeys)}\n    ],\n\n`;
    if (greenKeys.size > 0) templateCode += `    greenKeys: [\n      ${formatKeys(greenKeys)}\n    ],\n\n`;
    if (redKeys.size > 0) templateCode += `    redKeys: [\n      ${formatKeys(redKeys)}\n    ],\n\n`;
    if (linkKeys.size > 0) templateCode += `    linkKeys: [\n      ${formatKeys(linkKeys)}\n    ],\n\n`;

    templateCode += `  };\n}\n${registryCode}${prescripCode}`;

    body.appendPageBreak();
    body.appendParagraph("/* --- REVERSE GENERATED CODE --- */").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(templateCode);
    }

  // Helpers 
    function cleanSpaces(str) { 
    return str ? String(str).replace(/[\u00a0\s]+/g, " ").trim() : ""; 
    }

    function getReverseRegistryMap() {
    const reverseMap = { textToKey: {}, urlToKey: {} };
    if (typeof FORMAT_REGISTRY === 'undefined') return reverseMap;
    for (const key in FORMAT_REGISTRY) {
    const val = FORMAT_REGISTRY[key];
    if (typeof val === 'string') {
    reverseMap.textToKey[cleanSpaces(val).toLowerCase()] = key;
    } else if (val && typeof val === 'object') {
    if (val.text) reverseMap.textToKey[cleanSpaces(val.text).toLowerCase()] = key;
    if (val.url) reverseMap.urlToKey[String(val.url).trim()] = key;
    }
    }
    return reverseMap;
    }

    function extractFormattingSpans(textElement, keys, reverseMap, newLinksRegistry) {
    const text = textElement.getText() || "";
    let start = 0;
    while (start < text.length) {
    const isBold = textElement.isBold(start), isUnderline = textElement.isUnderline(start),
    isItalic = textElement.isItalic(start), color = textElement.getForegroundColor(start),
    linkUrl = textElement.getLinkUrl(start);
    let end = start;
    while (end < text.length && textElement.isBold(end) === isBold && textElement.isUnderline(end) === isUnderline &&
    textElement.isItalic(end) === isItalic && textElement.getForegroundColor(end) === color && textElement.getLinkUrl(end) === linkUrl) {
    end++;
    }
    let rawSpan = text.substring(start, end);
    let spanClean = cleanSpaces(rawSpan);

    if (spanClean.length > 1) {
    // Check if this text or URL is ALREADY in the registry
    let existingKey = reverseMap.textToKey[spanClean.toLowerCase()] || (linkUrl ? reverseMap.urlToKey[linkUrl] : null);

    if (linkUrl) {
    let linkKey = existingKey;
    if (!linkKey) {
    // Truly new link: create key and add to new registry
    linkKey = spanClean.replace(/[^A-Za-z0-9]/g, "_").toUpperCase() + "_ARTICLE";
    newLinksRegistry[linkKey] = { text: spanClean, url: linkUrl };
    }
    keys.linkKeys.add(linkKey);
    } else {
    const isGreen = (color === "#008000" || color === "#b6d7a8"), 
      isRed = (color === "#ff0000" || color === "#ea9999");

    let identifier = existingKey || spanClean;

    if (isGreen) keys.greenKeys.add(identifier);
    else if (isRed) keys.redKeys.add(identifier);
    else {
    if (isBold && isUnderline) keys.boldUnderlineKeys.add(identifier);
    else if (isBold) keys.boldKeys.add(identifier);
    if (isItalic) keys.italicKeys.add(identifier);
    }
    }
    }
    start = end;
    }
    }

    function convertToGPronouns(text) {
    if (!text) return "";
    const pronouns = { 
    " he ": " ${g.he} ", " she ": " ${g.he} ", 
    " him ": " ${g.him} ", " her ": " ${g.him} ", 
    " his ": " ${g.his} ", " hers ": " ${g.his} ", 
    " He ": " ${g.he} ", " She ": " ${g.he} ",
    " His ": " ${g.his} ", " Her ": " ${g.his} "
    };
    let newText = text;
    for (const [key, val] of Object.entries(pronouns)) { 
    newText = newText.replace(new RegExp(key, "g"), val); 
    }
    return newText;
    }

    function escapeBackticks(text) { 
    return String(text || "").replace(/`/g, "\\`").replace(/\$/g, "\\$"); 
    }

/* ------------------ CANINE WELLNESS ------------------ */
  // 8 Week Wellness Template
    function generate8WkWellnessTemplate(size, sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);

    const spayorneuterText = sex === 'female'
    ? `Spay: It is recommended you have your ${g.dog} spayed at 6 months if you do not intend to breed ${g.him}. While it isn’t wrong to keep ${g.him} intact, intact females are at risk of developing several life threatening diseases such as breast cancer, diabetes, & pyometra. 1 in 4 female dogs will get breast cancer if they are not spayed by their second heat cycle. In dogs there is a 50% chance breast cancer spreads throughout the body. The earliest time to have your ${g.dog} spayed is when ${g.he} ${g.is} 6 months old before ${g.his} first heat cycle. This will give ${g.his} body enough time to grow while also reducing the risk of diseases that are associated with spaying too early. Even if ${g.he} ${g.has} already had ${g.his} second heat cycle, spaying is still recommended since many mammary tumors are stimulated by estrogen & pyometra is still a possibility.`
    : size === 'large'
    ? `Neuter: On physical exam I was able to identify both of your ${g.dogs} testicles in ${g.his} scrotum. It is recommended you have ${g.him} neutered once ${g.he} ${g.is} 10 - 12 months old if you don’t intend to breed ${g.him}. By 10 - 12 months of age, large breed dogs such as ${g.him} have already received all the testosterone they need in order to grow normally. While it isn’t wrong to keep ${g.him} intact, neutering ${g.him} reduces or completely eradicates the risk of certain diseases such as prostatitis, several types of cancer, & hernias to name a few.`
    : `Neuter: On physical exam I was able to identify both of your ${g.dogs} testicles in ${g.his} scrotum. It is recommended you have ${g.him} neutered once ${g.he} ${g.is} 6 months old if you don’t intend to breed ${g.him}. By 6 months of age, smaller breed dogs such as ${g.him} have already received all the testosterone they need in order to grow normally. While it isn’t wrong to keep ${g.him} intact, neutering ${g.him} reduces or completely eradicates the risk of certain diseases such as prostatitis, several types of cancer, & hernias to name a few.`;

    let dentalProducts;
    if (sex === 'female') {
    dentalProducts = ['small dog toothbrush', 'medium/large dog toothbrush', 'animal safe toothpaste'];
    } else if (size === 'large') {
    dentalProducts = ['medium/large dog toothbrush', 'animal safe toothpaste'];
    } else {
    dentalProducts = ['small dog toothbrush', 'animal safe toothpaste'];
    }

    // Join the list with "and" properly
    const dentalListText = dentalProducts.length > 2 
    ? `${dentalProducts.slice(0, -1).join(', ')}, & ${dentalProducts[dentalProducts.length - 1]}`
    : dentalProducts.join(' & ');

    const text = [
    `Vaccines: Your ${g.dog} ${g.has} received ${g.his} first ${g.round} of ${g.puppy} vaccines today. Because of the antibodies that ${g.he} received from ${g.his} ${g.mother} milk, the vaccines won’t provide full immunity until 16 weeks of age when most of the ${g.mother} antibodies have disappeared. For that reason, it’s important to booster them every 3 - 4 weeks as your ${g.dogs} immune system slowly takes over.`,
    `Because your ${g.dog} ${g.is} still getting vaccines to better ${g.his} immune system, it’s best to keep ${g.him} away from dog parks & other dogs that aren’t part of your household until two weeks after ${g.he} ${g.has} finished ${g.his} puppy vaccines.`,
    `The initial distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo ${g.shot} in the left hindlimb. The 1 year bordetella vaccine was given orally. You may notice that after vaccination your ${g.dog} ${g.is} more tired than usual, ${g.eats} less, or ${g.is} sore at the injection ${g.site}, & this is perfectly normal.`,
    `Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your ${g.dog} back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your ${g.dog}.`,
    `Heartworms: Heartworms are spread by mosquitoes which don’t die in the Texas "winter", so our pets are at risk of infection year round. Furthermore, heartworms can be fatal & there is a risk of death even with proper treatment. Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your ${g.dog} on monthly preventatives such as Heartgard, Simparica Trio, Revolution, etc. Depending on the brand, they can protect your ${g.dog} from heartworms, fleas, ticks, & common intestinal parasites with a single treatment. These can be given orally or topically & are generally well tolerated. Because your ${g.dog} ${g.is} still growing, you will need to come back once a month to have ${g.him} weighed & get the appropriate dose of preventative.`,
    `${spayorneuterText}`,
    `Food: A high quality diet is the best way to keep your ${g.dog} healthy. Any puppy diet from Hill’s Science Diet (Hill's puppy dry food or Hill's puppy wet food), Purina Pro Plan (Purina puppy dry food or Purina puppy wet food), or Royal Canin (RC puppy dry food or RC puppy wet food) are all acceptable. A puppy diet is advised until your ${g.dog} ${g.is} a year old at which point you can transition to an adult diet. Dry food & wet food are both appropriate to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag/can for a dog of ${g.his} weight.`,
    `Dental care: The best way to keep your ${g.dogs} teeth healthy is to brush them daily for 10 seconds total using a ${dentalListText}. Animal safe toothpaste such as C.E.T. can be purchased from the clinic or from online stores. Getting your ${g.dog} used to having ${g.his} teeth brushed early will improve ${g.his} overall health.`,
    `You can start by having ${g.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then applying the pet safe toothpaste & letting ${g.him} lick it off every day for a week. Finally, gently brush ${g.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
    `If your ${g.dog} resists having ${g.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${g.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`,
    `Next appointment: Bring your ${g.dog} back in 3 - 4 weeks for ${g.his} next ${g.round} of ${g.puppy} vaccines.`
    ].join('\n');

    return {
    sex,
    plurality,
    text,
    boldKeys: [
    'VACCINES_HEADER',
    'HEARTWORMS_HEADER',
    'NEUTER_HEADER',
    'SPAY_HEADER',
    'DIET_HEADER',
    'DENTAL_HEADER',
    'NEXT_APPOINTMENT_HEADER',
    'DAPP_INITIAL',
    'BORDETELLA_VXN',
    ],

    boldUnderlineKeys: [
    'PUPPY_QUARANTINE',
    'VXN_RXN',
    'IMMEDIATELY',
    'RARE_RXN',
    'SMALL_NEUTER',
    'LARGE_NEUTER',
    'SPAY1',
    'SPAY2',
    'SPAY3',
    'GRAIN_FREE',
    'DENTAL_BRUSHING',
    'DENTAL_BRUSHING_CORE',
    'XYLITOL',
    ],

    linkKeys: [
    'VOHC_DOG_LINK',
    'SMALL_DOG_TOOTHBRUSH_LINK',
    'LARGE_TOOTHBRUSH_LINK',
    'TOOTHPASTE_LINK',
    'HILLS_PUPPY_DRY_LINK',
    'HILLS_PUPPY_WET_LINK',
    'PURINA_PUPPY_DRY_LINK',
    'PURINA_PUPPY_WET_LINK',
    'ROYAL_CANIN_PUPPY_DRY_LINK',
    'ROYAL_CANIN_PUPPY_WET_LINK',
    ],
    };
    }

  // 12 Week Wellness Template
    function generate12WkWellnessTemplate(size, sex, plurality = 'singular') {
    // 1. Initialize the Grammar Helper
    const g = getGrammar('wellness', plurality, sex);

    // 2. Start from the 8-week template (passing plurality through)
    const template = generate8WkWellnessTemplate(size, sex, plurality);

    // 3. New vaccine text using the grammar object
    const newVaccineText12wks = [
        `Vaccines: Your ${g.dog} ${g.has} received ${g.his} next ${g.round} of ${g.puppy} vaccines today. Because of the antibodies that ${g.he} received from ${g.his} ${g.mother} milk, ${g.he} won’t have full immunity until 2 weeks after ${g.his} last ${g.round} of ${g.puppy} vaccines. Keep ${g.him} away from dog parks, training facilities, and other dogs until then.`,
        `The distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo ${g.shot} with the initial lepto vaccine in the left hindlimb. You may notice that after vaccination your ${g.dog} ${g.is} more tired than usual, ${g.eats} less, or ${g.is} sore at the injection site, & this is perfectly normal.`,
        `Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your ${g.dog} back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your ${g.dog}.`
    ].join('\n');

    // 4. Update the text replacement logic
    // We use a broader regex so it finds the block regardless of "dog" vs "dogs"
    template.text = template.text.replace(
        /Vaccines:[\s\S]*?not expected to occur in your .*?\./,
        newVaccineText12wks
    );

    // 5. Replace the next appointment sentence
    template.text = template.text.replace(
        /Next appointment:[\s\S]*?\./,
        `Next appointment: Bring your ${g.dog} back in 4 weeks for ${g.his} final ${g.round} of ${g.puppy} vaccines.`
    );

    // 6. Update Keys
    template.boldKeys = [
        ...(template.boldKeys || []),
        'DAPP_BOOSTER',
        'LEPTO_INITIAL',
    ];

    template.boldUnderlineKeys = [
        ...(template.boldUnderlineKeys || []),
        'Quarantine_12WK',
    ];

    return template;
    }

  // 16 Week Wellness Template
    function generate16WkWellnessTemplate(size, sex, plurality = 'singular') {
    // 1. Initialize the Grammar Helper
    const g = getGrammar('wellness', plurality, sex);

    // 2. Start from the 8-week template
    const template = generate8WkWellnessTemplate(size, sex, plurality);

    // 3. New vaccine text using the grammar object
    const newVaccineText16wks = [
        `Vaccines: Your ${g.dog} ${g.has} received ${g.his} final ${g.round} of ${g.puppy} vaccines today. The antibodies ${g.he} received from ${g.his} ${g.mother} milk have mostly decreased at this point, so ${g.his} own immune system should be in full effect. Starting from today, ${g.he} can get ${g.his} vaccines once a year.`,
        `Over the next two weeks your ${g.dog} will build up immunity to the vaccines given. During this time, keep ${g.him} away from dog parks & other dogs that aren’t part of your household. Afterwards ${g.he} ${g.is} safe to interact with other dogs.`,
        `The 1 year rabies vaccine was given in the right hindlimb. The 1 year distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo ${g.shot} with the 1 year lepto vaccine in the left hindlimb. You may notice that after vaccination your ${g.dog} ${g.is} more tired than usual, ${g.eats} less, or ${g.is} sore at the injection ${g.site}, & this is perfectly normal.`,
        `Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your ${g.dog} back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your ${g.dog}.`
    ].join('\n');

    // 4. Replace the vaccine block
    // Using a flexible regex to catch the block regardless of "dog" vs "dogs"
    template.text = template.text.replace(
        /Vaccines:[\s\S]*?not expected to occur in your .*?\./,
        newVaccineText16wks
    );

    // 5. Replace Next Appointment with dynamic spay/neuter terms
    const procedure = sex === 'female' ? 'spay' : 'neuter';
    
    template.text = template.text.replace(
        /Next appointment:[\s\S]*$/,
        `Next appointment: Bring your ${g.dog} back once ${g.he} ${g.is} at least 6 months old for ${g.his} ${procedure} & heartworm test. If you would not like to ${procedure} ${g.him}, bring ${g.him} back in 1 year for ${g.his} annual adult vaccines.`
    );

    // 6. Update Keys
    template.boldKeys = [
        ...(template.boldKeys || []),
        'RABIES_1YR',
        'DAPP_1YR',
        'LEPTO_VXN',
    ];

    template.boldUnderlineKeys = [
        ...(template.boldUnderlineKeys || []),
        'Quarantine_16WK',
    ];

    return template;
    }

  // Initial Adult Vaccine Template
    function generateInitialAdultTemplate(sex, plurality = 'singular', size) {
    const g = getGrammar('wellness', plurality, sex);

    // Dynamic Dental Products List
    let dentalProducts;
    if (sex === 'female') {
        dentalProducts = ['small dog toothbrush', 'medium/large dog toothbrush', 'animal safe toothpaste'];
    } else if (size === 'large') {
        dentalProducts = ['medium/large dog toothbrush', 'animal safe toothpaste'];
    } else {
        dentalProducts = ['small dog toothbrush', 'animal safe toothpaste'];
    }

    const dentalListText = dentalProducts.length > 2 
        ? `${dentalProducts.slice(0, -1).join(', ')}, & ${dentalProducts[dentalProducts.length - 1]}`
        : dentalProducts.join(' & ');

    const text = [
        `Vaccines: Your ${g.dog} ${g.has} received ${g.his} first ${g.round} of adult vaccinations. The 1 year rabies vaccine was given in the right hindlimb. The initial distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo ${g.shot} with the initial lepto vaccine in the left hindlimb. The 1 year bordetella vaccine was given orally. Your ${g.dog} will need a booster of the DAPP and lepto vaccines in 3 - 4 weeks.`,
        
        `You may notice that after vaccination your ${g.dog} ${g.is} more tired than usual, ${g.eats} less, or ${g.is} sore at the injection sites, & this is perfectly normal. Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your ${g.dog} back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your ${g.dog}.`,
        
        `Heartworms: A heartworm test was performed on your ${g.dog}. We will contact you in 3 - 4 business days with the results. Heartworms are spread by mosquitoes which don’t die in the Texas "winter", so our pets are at risk of infection year round. Furthermore, heartworms can be fatal & there is a risk of death even with proper treatment. Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your ${g.dog} on monthly preventatives such as Heartgard, Nexgard, Simparica Trio, Revolution, etc.`,
        
        `Early detection labwork: Yearly blood work is recommended for ${g.dogs} the same as it is in humans and starts at 3 years of age. This lets us get a baseline for your pet and allows us to catch abnormalities before they’re noticeable outwardly. Depending on the panel run, this can check for issues in the liver, kidneys, thyroid, bladder, glucose, and many other organs and values. At 6 years of age, a larger panel for “senior” pets is advised.`,
        
        `Food: A high quality diet is the best way to keep your ${g.dog} healthy. Food from Hill’s Science Diet (Hill's dog dry food or Hill's dog wet food), Purina Pro Plan (Purina dog dry food or Purina dog wet food), or Royal Canin (RC dog dry food or RC dog wet food) are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in ${g.dogs}, so either are wonderful to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag/can for a dog of ${g.his} weight.`,

        `Dental care: The best way to keep your ${g.dogs} teeth healthy is to brush them daily for 10 seconds total using a ${dentalListText}. Animal safe toothpaste such as C.E.T. can be purchased from the clinic or from online stores. Getting your ${g.dog} used to having ${g.his} teeth brushed early will improve ${g.his} overall health.`,
        
        `You can start by having ${g.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then applying the pet safe toothpaste & letting ${g.him} lick it off every day for a week. Finally, gently brush ${g.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
        
        `If your ${g.dog} resists having ${g.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${g.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`,
        
        `Next appointment: Bring your ${g.dog} back in 3 - 4 weeks for a booster of ${g.his} vaccines.`
    ].join('\n');

    return {
    sex,
    text,
    boldKeys: [
    'VACCINES_HEADER',
    'HEARTWORMS_HEADER',
    'DIET_HEADER',
    'DENTAL_HEADER',
    'NEXT_APPOINTMENT_HEADER',
    'RABIES_1YR',
    'DAPP_INITIAL',
    'LEPTO_INITIAL',
    'BORDETELLA_VXN',
    'LABWORK',
    ],

    boldUnderlineKeys: [
    'VXN_BOOSTER',
    'VXN_RXN',
    'IMMEDIATELY',
    'RARE_RXN',
    'HEARTWORM_TEST',
    'LAB_RESULTS',
    'GRAIN_FREE',
    'DENTAL_BRUSHING',
    'DENTAL_BRUSHING_CORE',
    'XYLITOL',
    ],

    linkKeys: [
    'VOHC_DOG_LINK',
    'SMALL_DOG_TOOTHBRUSH_LINK',
    'LARGE_TOOTHBRUSH_LINK',
    'TOOTHPASTE_LINK',
    'HILLS_DOG_DRY_LINK',
    'HILLS_DOG_WET_LINK',
    'PURINA_DOG_DRY_LINK',
    'PURINA_DOG_WET_LINK',
    'ROYAL_CANIN_DOG_DRY_LINK',
    'ROYAL_CANIN_DOG_WET_LINK',
    ],
    };
    }

  // 1-Year Adult Vaccine Template
    function generate1YearAdultTemplate(sex, plurality = 'singular', size) {
    // 1. Initialize the Grammar Helper
    const g = getGrammar('wellness', plurality, sex);

    const dentalProducts = [
    'small dog toothbrush',
    'medium/large dog toothbrush',
    'animal safe toothpaste'
    ];

    // 3. Main Template Text
      const text = [ 
      `Vaccines: Your ${g.dog} ${g.has} received ${g.his} first ${g.round} of adult vaccinations. Because you have kept to ${g.his} vaccination schedule, ${g.his} immune system will not need another booster until next year.`,
      `The 1 year rabies vaccine was given in the right hindlimb. The 1 year distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo ${g.shot} with the 1 year lepto vaccine in the left hindlimb. The 1 year bordetella vaccine was given orally. You may notice that after vaccination your ${g.dog} ${g.is} more tired than usual, ${g.eats} less, or ${g.is} sore at the injection sites, & this is perfectly normal.`,
      `Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your ${g.dog} back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your ${g.dog}.`,
      `Heartworms: A heartworm test was performed on your ${g.dog}. We will contact you in 3 - 4 business days with the results. Heartworms are spread by mosquitoes which don’t die in the Texas "winter", so our pets are at risk of infection year round. Furthermore, heartworms can be fatal & there is a risk of death even with proper treatment. Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your ${g.dog} on monthly preventatives such as Heartgard, Nexgard, Simparica Trio, Revolution, etc.`,
      `Early detection labwork: Samples were drawn from your ${g.dog}. You will receive a call in 3 - 4 business days with the results. Yearly blood work is recommended for dogs the same as it is in humans for the sake of monitoring for abnormalities that aren’t visible from the outside. Depending on the panel run, this can check for issues in the liver, kidneys, thyroid, bladder, glucose, and many other organs and values. If no abnormalities are found, the results can be used as a baseline so that your ${g.dogs} overall health is closely monitored.`,
      `Food: A high quality diet is the best way to keep your ${g.dog} healthy. If you haven’t already, you can transition ${g.him} from ${g.his} ${g.puppy} diet to ${g.his} adult diet. Food from Hill’s Science Diet (Hill's dog dry food or Hill's dog wet food), Purina Pro Plan (Purina dog dry food or Purina dog wet food), or Royal Canin (RC dog dry food or RC dog wet food) are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in ${g.dogs}, so either are wonderful to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag/can for a dog of ${g.his} weight.`,
      `Dental care: The best way to keep your ${g.dogs} teeth healthy is to brush them daily for 10 seconds total using a ${dentalProducts.slice(0, -1).join(', ')} & ${dentalProducts[dentalProducts.length - 1]}. Animal safe toothpaste such as C.E.T. can be purchased from the clinic or from online stores. Getting your ${g.dog} used to having ${g.his} teeth brushed early will improve ${g.his} overall health.`,
      `You can start by having ${g.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then applying the pet safe toothpaste & letting ${g.him} lick it off every day for a week. Finally, gently brush ${g.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
      `If your ${g.dog} resists having ${g.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${g.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`,
      `Next appointment: Bring your ${g.dog} back one year from today for ${g.his} next annual vaccines.`
      ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["WELLNESS"],
    boldKeys: [
    'VACCINES_HEADER',
    'HEARTWORMS_HEADER',
    'DIET_HEADER',
    'DENTAL_HEADER',
    'NEXT_APPOINTMENT_HEADER',
    'RABIES_1YR',
    'DAPP_1YR',
    'LEPTO_VXN',
    'BORDETELLA_VXN',
    'LABWORK',
    ],

    boldUnderlineKeys: [
    'VXN_RXN',
    'IMMEDIATELY',
    'RARE_RXN',
    'HEARTWORM_TEST',
    'LAB_RESULTS',
    'GRAIN_FREE',
    'DENTAL_BRUSHING',
    'DENTAL_BRUSHING_CORE',
    'XYLITOL',
    ],

    linkKeys: [
    'VOHC_DOG_LINK',
    'SMALL_DOG_TOOTHBRUSH_LINK',
    'LARGE_TOOTHBRUSH_LINK',
    'TOOTHPASTE_LINK',
    'HILLS_DOG_DRY_LINK',
    'HILLS_DOG_WET_LINK',
    'PURINA_DOG_DRY_LINK',
    'PURINA_DOG_WET_LINK',
    'ROYAL_CANIN_DOG_DRY_LINK',
    'ROYAL_CANIN_DOG_WET_LINK',
    ],
    };
    }

  // 2-Year Adult Vaccine Template
    function generate2YearAdultTemplate(sex, plurality = 'singular',size) {
    // 1. Initialize Grammar
    const g = getGrammar('wellness', plurality, sex);
    
    // 2. Start from the upgraded 1-year template
    const template = generate1YearAdultTemplate(sex, plurality, size);

    // 3. Update Bold Keys to include 3-year versions
    template.boldKeys = [
        ...(template.boldKeys || []),
        'RABIES_3YR',
        'DAPP_3YR',
    ];

    // 4. Replace the 1-year text with 3-year text
    // We use flexible regex (.*? or [g.is]) to account for singular/plural differences
    template.text = template.text.replace(
        /Your .*? (has|have) received .*? first .*? of adult vaccinations\./,
        `Your ${g.dog} ${g.has} received ${g.his} annual adult vaccinations.`
    );

    template.text = template.text
        .replace(
            /The 1 year rabies vaccine was given in the right hindlimb\./,
            'The 3 year rabies vaccine was given in the right hindlimb.'
        )
        .replace(
            /The 1 year distemper, adenovirus, parvovirus, & parainfluenza \(DAPP\) vaccine/,
            'The 3 year distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine'
        );

    // 5. Remove puppy food transition (using flexible regex for him/her/them)
    template.text = template.text.replace(
        /If you haven’t already, you can transition .*? from .*? puppy diet to .*? adult diet\.\s*/i,
        ''
    );

    return template;
    }

  // 2-Year Lepto Vaccine Template
    function generate2YearLeptoTemplate(sex, plurality = 'singular', size = 'small') {
    // 1. Pass parameters in the correct Sex-First order
    const template = generate2YearAdultTemplate(sex, plurality, size);

    // 2. Replace the vaccine listing paragraph ONLY
    template.text = template.text.replace(
        /The 3 year rabies vaccine was given in the right hindlimb\. The 3 year distemper, adenovirus, parvovirus, & parainfluenza \(DAPP\) vaccine was given as a combo .*? with the 1 year lepto vaccine in the left hindlimb\. The 1 year bordetella vaccine was given orally\./,
        'The 1 year lepto vaccine was given in the left hindlimb. The 1 year bordetella vaccine was given orally.'
    );

    // 3. Update Bold Keys
    template.boldKeys = template.boldKeys.filter(key => 
        key !== 'RABIES_3YR' && key !== 'DAPP_3YR'
    );

    return template;
    }

  // 7-Year Adult Vaccine Template                                                             
    function generate7YearAdultTemplate(sex, plurality = 'singular',size) {
    // 1. Pass parameters to the 2-year base (which handles the 3-yr vaccines)
    const template = generate2YearAdultTemplate(sex, plurality, size);
    const g = getGrammar('wellness', plurality, sex);

    // 2. Adjust senior diet wording
    // The regex is widened to handle "dog" vs "dogs" and "his/her" vs "their"
    template.text = template.text.replace(
        /Food: A high quality diet is the best way to keep your .*? healthy\.[\s\S]*?can for a .*? of .*? weight\./,
        `Food: A high quality diet is the best way to keep your ${g.dog} healthy. ${g.Dogs} that are older than 7 years are advised to be on a senior diet. Food from Hill’s Science Diet (Hill's senior dog dry food or Hill's senior dog wet food), Purina Pro Plan (Purina senior dog dry food or Purina senior dog wet food), or Royal Canin (RC senior dog dry food or RC senior dog wet food) are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in dogs, so either are wonderful to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag/can for a dog of ${g.his} weight.`
    );

    // 3. Add senior dog links to the existing link keys
    template.linkKeys = [
        ...(template.linkKeys || []),
        'HILLS_SR_DOG_DRY_LINK',
        'HILLS_SR_DOG_WET_LINK',
        'PURINA_SR_DOG_DRY_LINK',
        'PURINA_SR_DOG_WET_LINK',
        'ROYAL_CANIN_SR_DOG_DRY_LINK',
        'ROYAL_CANIN_SR_DOG_WET_LINK',
    ];

    return template;
    }

  // 7-Year Lepto Vaccine Template
    function generate7YearLeptoTemplate(sex, plurality = 'singular', size = 'small') {
    // 1. Pass parameters in the established Sex-First order
    const template = generate2YearAdultTemplate(sex, plurality, size);
    const g = getGrammar('wellness', plurality, sex);

    // 2. Replace the vaccine listing paragraph ONLY
    template.text = template.text.replace(
        /The 3 year rabies vaccine was given in the right hindlimb\. The 3 year distemper, adenovirus, parvovirus, & parainfluenza \(DAPP\) vaccine was given as a combo .*? with the 1 year lepto vaccine in the left hindlimb\. The 1 year bordetella vaccine was given orally\./,
        'The 1 year lepto vaccine was given in the left hindlimb. The 1 year bordetella vaccine was given orally.'
    );

    // 3. Adjust senior diet wording
    // Fixed the very last instance of "dog" to use ${g.dog}
    template.text = template.text.replace(
        /Food: A high quality diet is the best way to keep your .*? healthy\.[\s\S]*?can for a .*? of .*? weight\./,
        `Food: A high quality diet is the best way to keep your ${g.dog} healthy. ${g.Dogs} that are older than 7 years are advised to be on a senior diet. Food from Hill’s Science Diet (Hill's senior dog dry food or Hill's senior dog wet food), Purina Pro Plan (Purina senior dog dry food or Purina senior dog wet food), or Royal Canin (RC senior dog dry food or RC senior dog wet food) are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in dogs, so either are wonderful to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag/can for a dog of ${g.his} weight.`
    );

    // 4. Update Keys
    // Clean up bold keys since Rabies/DAPP were removed
    template.boldKeys = template.boldKeys.filter(key => 
        key !== 'RABIES_3YR' && key !== 'DAPP_3YR'
    );

    // Add senior dog links
    template.linkKeys = [
        ...(template.linkKeys || []),
        'HILLS_SR_DOG_DRY_LINK',
        'HILLS_SR_DOG_WET_LINK',
        'PURINA_SR_DOG_DRY_LINK',
        'PURINA_SR_DOG_WET_LINK',
        'ROYAL_CANIN_SR_DOG_DRY_LINK',
        'ROYAL_CANIN_SR_DOG_WET_LINK',
    ];

    return template;
    }

  // Canine Overweight | 1st
    function generateCanineOverweightTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);

    const text = [
        `Weight: Your ${g.dog} ${g.weighs} more than the average ${g.dog} of ${g.his} size. Ideally we would be able to feel ${g.his} ribs but not see them. Helping ${g.him} to lose weight can increase ${g.his} life span by as much as 1 ½ years. The best way to lose weight is through diet.`,
        
        `You can use the diet ${g.he} ${g.is} currently on or you can use a prescription weight loss food from Hill’s Prescription Diet (Hill’s weight loss dry food or Hill’s weight loss wet food), Purina Pro Plan (Purina weight loss dry food or Purina weight loss wet food), or Royal Canin (RC weight loss dry food or RC weight loss wet food). Regardless, begin by measuring how much your ${g.dog} ${g.eats} using a measuring cup. Make sure to feed twice daily on a schedule rather than leaving food down at all times. If ${g.he} ${g.steals} food from siblings, you may need to feed separately. Finally, decrease ${g.his} food by 10 - 25%.`,
        
        `We’re aiming to have ${g.him} lose 1 - 2% of ${g.his} body weight per week. If ${g.he} ${g.begins} losing more than that per week, increase the amount of food ${g.he} ${g.gets}. Another way you can help ${g.him} lose weight is by converting ${g.his} treats into healthy alternatives such as slices of apples, carrots, ice cubes, cucumbers, or green beans.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["OVERWEIGHT"],
    boldKeys: [
    'WEIGHT_HEADER',
    ],

    boldUnderlineKeys: [
    'DIET_LIFESPAN',
    'DIET_WEEKLY_GOAL',
    'OVERWEIGHT_WARNING',
    ],

    linkKeys: [
    'HILLS_DOG_DIET_DRY_LINK',
    'HILLS_DOG_DIET_WET_LINK',
    'PURINA_DOG_DIET_DRY_LINK',
    'PURINA_DOG_DIET_WET_LINK',
    'ROYAL_CANIN_DOG_DIET_DRY_LINK',
    'ROYAL_CANIN_DOG_DIET_WET_LINK',
    ],
    };
    }

  // Canine Overweight | 2nd, Continue
    function generateCanineOverweight2Template(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);

    const text = [
        `Weight: Congrats on helping your ${g.dog} lose weight! Continuing to help ${g.him} lose weight can extend ${g.his} life span by as much as 1 ½ years.`,
        
        `As a reminder, food from Hill’s Prescription Diet (Hill’s weight loss dry food or Hill’s weight loss wet food), Purina Pro Plan (Purina weight loss dry food or Purina weight loss wet food), or Royal Canin (RC weight loss dry food or RC weight loss wet food) can be used as needed. Otherwise, continue measuring how much ${g.he} ${g.eats} using a measuring cup and feeding on a twice daily schedule rather than leaving food down at all times. Separate ${g.him} from siblings at meal time if necessary.`,
        
        `We’re aiming to have ${g.him} lose 1 - 2% of ${g.his} body weight per week. If ${g.he} ${g.begins} losing more than that per week, increase the amount of food ${g.he} ${g.gets}. Another way you can help ${g.him} lose weight is by converting ${g.his} treats into healthy alternatives such as slices of apples, carrots, ice cubes, cucumbers, or green beans.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["OVERWEIGHT"],
    boldKeys: [
    'WEIGHT_HEADER',
    ],

    boldUnderlineKeys: [
    'DIET2_LIFESPAN',
    'DIET_WEEKLY_GOAL',
    ],

    greenKeys: [],
    linkKeys: [
    'HILLS_DOG_DIET_DRY_LINK',
    'HILLS_DOG_DIET_WET_LINK',
    'PURINA_DOG_DIET_DRY_LINK',
    'PURINA_DOG_DIET_WET_LINK',
    'ROYAL_CANIN_DOG_DIET_DRY_LINK',
    'ROYAL_CANIN_DOG_DIET_WET_LINK',
    ],
    };
    }

  // Canine Healthy Weight
    function generateDogHealthyWeightTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Weight: Your dog is a healthy weight for a dog of ${p.his} size. ${p.His} ribs can be felt without difficulty and ${p.he} has a slight waist. Keeping ${p.him} around ${p.his} current weight will help ${p.him} live approximately 1 ½ years longer than ${p.he} would if ${p.he} were over or underweight.`,
    `Continue to monitor ${p.his} weight and feed ${p.him} as you’ve been doing. Signs of an overweight dog include difficulty feeling the ribs and loss of a waist when viewed from above. You can switch ${p.his} treats to apple slices, carrots, green beans, ice cubes, or cucumbers if you notice ${p.him} starting to gain weight. Signs of an underweight dog are the spine being visible in the same fashion as your knuckles, ribs visible enough to be counted, and hips that can be felt when running your hand over your dog’s back end.`
    ].join('\n');

    return {
    sex,
    text,
    boldKeys: [
    'WEIGHT_HEADER',
    ],

    boldUnderlineKeys: [
    'HEALTHY_WEIGHT_LIFESPAN',
    'OVERWEIGHT_DOG_SIGNS',
    'UNDERWEIGHT_DOG_SIGNS',
    ],

    greenKeys: [],
    linkKeys: [],
    };
    }
  
  // Canine Underweight
    function generateDogUnderweightTemplate(sex) {
    const p = getPronoun(sex);

    const text = [
    `Underweight: Your dog weighs less than the average dog of ${p.his} size. Ideally, we would be able to feel ${p.his} ribs but not see them. Helping ${p.him} gain weight can increase ${p.his} quality of life.`,
    `The best way for ${p.him} to gain weight is through diet. Food from Hill’s Science Diet, Purina Pro Plan, or Royal Canin are all wonderful diets as they’re formulated by veterinary scientists. You can also add lukewarm water to the food or low sodium chicken broth to increase the smell and flavor.`,
    `Increase how much ${p.he} eats by as much as 25 - 50%. We're aiming to have ${p.him} gain approximately 10% of ${p.his} current weight. Failure to gain weight is concerning for disease and would prompt us to perform tests such as labwork, ultrasound, or x-rays.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["UNDERWEIGHT"],

    boldKeys: [
    'UNDERWEIGHT_HEADER',
    ],

    boldUnderlineKeys: [
    'UNDERWEIGHT_LIFESPAN',
    `UNDERWEIGHT_GOAL`,
    'UNDERWEIGHT_WARNING',
    ],

    };
    }

  // Canine Vaccine Information
    function generateDogVaccineInformationTemplate() {
    const text = [
    `Rabies: Rabies is a fatal virus that is spread from wild animal bites to our pets & humans. The most common spreaders in Texas are raccoons, skunks, bats, foxes, & coyotes. Signs of rabies start with voice changes & becoming shy. Next the pet acts aggressive or becomes paralyzed. The animal then dies if it isn't euthanized by then.`,
    `THERE IS NO CURE FOR RABIES. The only way to test for rabies involves decapitating an animal & taking samples of the brain. State law requires any animal that is exposed to rabies to either undergo quarantine for up to 6 months or be euthanized.`,
    `For the sake of your pet’s health, get regular rabies vaccines. You can learn more about rabies from the Rabies in Animals article on Veterinary Partner or the Texas government website.`,
    `DAPP: Distemper, adenovirus, parvovirus, & parainfluenza virus (DAPP) are a series of viruses that cause serious, contagious disease in our dogs. Dogs infected with distemper may have fever, coughing, difficulty breathing, skin infection, blindness, abortions, or seizures. Some studies suggest that distemper can be spread to humans. There is no cure for distemper. Similarly, parvovirus often causes fatal diarrhea in puppies & spreads rapidly. It requires hospitalization to effectively cure. Adenovirus & parainfluenza can cause respiratory infection in dogs. You can read the distemper article, the parvovirus article, & the adenovirus & parainfluenza article on Veterinary Partner for more information.`,
    `Lepto: Leptospirosis is a bacteria spread in the waterways & anywhere that woodland creatures (squirrels, raccoons, etc.) urinate. It causes liver & kidney disease & requires hospitalization to treat. Lepto can quickly kill our pets & can be spread to humans. Vaccination is strongly recommended to prevent the disease. You can read more about lepto from the Leptospirosis in Dogs article on Veterinary Partner.`,
    `Bordetella: Bordetella, also known as kennel cough, is a bacteria that is spread through the air whenever a dog enters an area that an infected dog has been. This includes dog parks, grooming facilities, & veterinary clinics. Most cases of bordetella resolve on their own without treatment, but some dogs get complicated cases that cause severe lung infections including pneumonia. The bordetella vaccine can be given orally, intranasally, or as an injection and rarely has reactions. You can learn more about bordetella from the Kennel Cough in Dogs article on Veterinary Partner.`
    ].join('\n');

    return {
    text,
    diagnoses: ["PARTIALLY_VACCINATED"],
    boldKeys: [
    'RABIES_HEADER',
    'DAPP_HEADER',
    'LEPTO_HEADER',
    'BORDETELLA_HEADER',
    ],

    boldUnderlineKeys: [
    `DAPP_WARNING`,
    `LEPTO_WARNING`,
    ],

    italicKeys: [
    'BORDETELLA_NAME'
    ],

    redKeys: [
    'RABIES_CURE',
    ],

    linkKeys: [
    'ADENOVIRUS_LINK',
    'BORDETELLA_LINK',
    'DISTEMPER_LINK',
    'LEPTO_LINK',
    'PARVO_LINK',
    'RABIES_IN_ANIMALS_LINK',
    'TEXAS_RABIES_LINK',
    ]
    };
    }

/* ------------------ CANINE OPTHALMOLOGY------------------ */
  // Canine Blind | 0, Partial
    function generateDogBlind0PartialTemplate(sex) {
    const p = getPronoun(sex);
    return {
    sex,
    diagnoses: ["PARTIALLY_BLIND"],
    text: [
    `Blind: Your dog shows signs of being partially blind. A small amount of vision (enough to see shadows & shapes) is present, but not enough to read or drive. Keep your living space free of obstacles to prevent your dog from tripping or bumping into things by mistake. Avoid making changes to your living space as your dog has most likely memorized the layout and will be confused if things move around. You can also use a Halo harness or similar devices to prevent your pet from running into objects.`
    ].join('\n'),
    boldKeys: [
    "BLIND_HEADER"
    ],
    boldUnderlineKeys: [

    ],

    linkKeys: [
    "HALO_HARNESS_ARTICLE"
    ],
    };
    }

  // Canine Blind | 1, Diagnosed
    function generateDogBlind1Template(sex) {
    const p = getPronoun(sex);
    const text = [
    `Blind: Your dog shows signs of being completely blind. Make sure to keep your living space free of obstacles to prevent your dog from tripping or bumping into things by mistake. Avoid making changes to your living space as your dog has most likely memorized the layout and will be confused if things move around. You can also use a Halo harness or similar devices to prevent your pet from running into objects.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["BLIND"],
    boldKeys: [
    "BLIND_HEADER"
    ],
    boldUnderlineKeys: [
    "BLIND_OBSTACLE_COURSE",
    ],

    linkKeys: [
    "HALO_HARNESS_ARTICLE",
    ],
    };
    }

  // Canine Blind | 2, Known
    function generateDogBlind2KnownTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Blind: Your dog is known to be completely blind. Make sure to keep your living space free of obstacles to prevent your dog from tripping or bumping into things by mistake. Avoid making changes to your living space as your dog has most likely memorized the layout and will be confused if things move around. You can also use a Halo harness or similar devices to prevent your pet from running into objects.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["BLIND"],
    boldKeys: [
    "BLIND_HEADER"
    ],

    boldUnderlineKeys: [
    "BLIND_OBSTACLE_COURSE"
    ],

    linkKeys: [
    "HALO_HARNESS_ARTICLE",
    ],
    };
    }

  // Canine Cherry Eye Template
    function generateDogCherryEyeTemplate(sex, plurality = 'singular') {
    const g = getGrammar('cherry_eye', plurality, sex);
    return {
    sex,
    diagnoses: ["CHERRY_EYE"],
    text: [
    `Cherry ${g.eye}: Your dog has a cherry ${g.eye}. This means that the ${g.gland} of the ${g.eye} that ${g.make} most of the tears ${g.is} poking out of ${g.its} normal position. Over time ${g.this} ${g.gland} can dry up & produce less tears, leading to a disorder known as dry eye. For dogs older than 1 year it is best to have the cherry ${g.eye} corrected as soon as possible. Surgery involves tucking the gland back in its normal position & using suture to prevent it from popping out again. You can learn more about cherry eyes from the Cherry Eye in Dogs and Cats article on Veterinary Partner.`
    ].join('\n'),
    diagnoses: ["CHERRY_EYE"],

    boldKeys: [
    "CHERRY_EYE_HEADER",
    ],

    boldUnderlineKeys: [
    "CHERRY_EYE_SURGERY_RECOMMENDATION",
    ],

    linkKeys: [
    "CHERRY_EYE_IN_DOGS_AND_CATS_ARTICLE"
    ],
    };
    }

/* ------------------ CANINE CARDIOLOGY ------------------ */
  // Canine Heart Murmur | 0th Discovered, No Tests
    function generateDogHeartMurmur0Template(sex) {
    const p = getPronoun(sex);
    const text = [
    `Heart murmur: A heart murmur was heard in your dog today. Heart murmurs are sounds produced whenever blood moves in a direction or location it isn’t meant to. Common causes include heartworms, heart disease, or fetal abnormalities. Grading is based on how loud the sound is. A higher grade (5 & 6) does not always indicate worse disease & a lower grade (1 & 2) does not always indicate a better disease. Diagnosis involves X-rays to see the shape & size of the heart can be performed in clinic and an echocardiogram to look at the inner workings of the heart and find a cause of disease can be scheduled as well. `,
    `In the meantime, monitor your dog for symptoms such as coughing, increased exhaustion when exercising, & low energy. Most importantly, count how fast your dog breathes while sleeping. If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease. You can learn more about heart murmurs from the Heart Murmurs in Dogs and Cats article on Veterinary Partner.`
    ].join('\n');


    return {
    sex,
    text,
    diagnoses: ["HEART_MURMUR"],
    boldKeys: [
    "HEART_MURMUR_HEADER"
    ],

    boldUnderlineKeys: [
    "HEART_MURMUR_HEARD",
    "HEART_MURMUR_GRADING",
    "HEART_MURMUR_WARNING_SIGNS"
    ],

    greenKeys: [
    "COMMON_CAUSES",
    "DIAGNOSIS",
    "SYMPTOMS",
    "TREATMENT",
    ],

    linkKeys: [
    "HEART_MURMUR_ARTICLE",
    ],
    };
    }

  // Canine Heart Murmur | 1st, Normal Radiographs
    function generateDogHeartMurmur1RadiographsNormalTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Heart murmur: A heart murmur was heard in your dog today. Heart murmurs are sounds produced whenever blood moves in a direction or location it isn’t meant to. Common causes include heartworms, heart disease, or fetal abnormalities. Grading is based on how loud the sound is. A higher grade (5 & 6) does not always indicate worse disease & a lower grade (1 & 2) does not always indicate a better disease.`,
    `X-rays to see the shape & size of the heart were performed and your dog’s heart doesn’t appear to be concerningly enlarged. At this time treatment with medication is not warranted, but an echocardiogram to look at the inner workings of the heart and diagnose the cause of the disease will need to be scheduled.`,
    `In the meantime, monitor your dog for symptoms such as coughing, increased exhaustion when exercising, & low energy. Most importantly, count how fast your dog breathes while sleeping. If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease. You can learn more about heart murmurs from the Heart Murmurs in Dogs and Cats article on Veterinary Partner.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["HEART_MURMUR"],
    boldKeys: [
    "HEART_MURMUR_HEADER" 
    ],

    boldUnderlineKeys: [
    "HEART_MURMUR_HEARD",
    "HEART_MURMUR_GRADING",
    "HEART_MURMUR_WARNING_SIGNS",
    "SCHEDULE_ECHO", 
    ],

    greenKeys: [
    "COMMON_CAUSES",
    "DIAGNOSE",
    "SYMPTOMS",
    "TREATMENT",
    ],

    linkKeys: [
    "HEART_MURMUR_ARTICLE",
    ]
    };
    }

  // Canine Heart Murmur | 1st, Cardiomegaly, Start Pimobendan
    function generateDogHeartMurmur1CardiomegalyTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Heart murmur: A heart murmur was heard in your dog today. Heart murmurs are sounds produced whenever blood moves in a direction or location it isn’t meant to. Common causes include heartworms, heart disease, or fetal abnormalities. Grading is based on how loud the sound is. A higher grade (5 & 6) does not always indicate worse disease & a lower grade (1 & 2) does not always indicate a better disease.`,
    `Diagnosis includes the x-rays that we performed to see the shape & size of the heart. These x-rays show that the heart is enlarged. It is compressing the lungs and trachea to an extent, so your dog will be started on medication to help improve heart function and slow the progression of disease. An echocardiogram to look at the inner workings of the heart and find a cause of disease will need to be scheduled.`,
    `In the meantime, give the medicine prescribed below as treatment to manage the disease. Monitor your dog for symptoms such as coughing, increased exhaustion when exercising, & low energy. Most importantly, count how fast your dog breathes while sleeping. If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease. You can learn more about heart murmurs from the Heart Murmurs in Dogs and Cats article on Veterinary Partner.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["HEART_MURMUR"],
    boldKeys: [
    "HEART_MURMUR_HEADER"
    ],

    boldUnderlineKeys: [
    "DIAGNOSIS_RESET",
    "HEART_MURMUR_GRADING",
    "HEART_MURMUR_HEARD",
    "HEART_MURMUR_WARNING_SIGNS",
    "SCHEDULE_ECHO"
    ],

    greenKeys: [
    "COMMON_CAUSES",
    "DIAGNOSIS",
    "SYMPTOMS",
    "TREATMENT",
    ],

    linkKeys: [
    "HEART_MURMUR_ARTICLE"
    ],
    };
    }

  // Canine Heart Murmur | 3rd, Known
    function generateDogHeartMurmur3KnownTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Heart murmur: Your dog is known to have a heart murmur. Heart murmurs are sounds produced whenever blood moves in a direction or location it isn’t meant to. Monitor your dog for symptoms such as coughing, increased exhaustion when exercising, & low energy. Most importantly, count how fast your dog breathes while sleeping. If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease. You can learn more about heart murmurs from the Heart Murmurs in Dogs and Cats article on Veterinary Partner.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["HEART_MURMUR"],
    boldKeys: [
    "HEART_MURMUR_HEADER"
    ],

    boldUnderlineKeys: [
    "HEART_MURMUR_WARNING_SIGNS",
    ],

    greenKeys: [
    "COMMON_CAUSES",
    "DIAGNOSIS",
    "SYMPTOMS",
    "TREATMENT",
    ],

    linkKeys: [
    "HEART_MURMUR_ARTICLE",
    ]
    };
    }

/* ------------------ CANINE GASTROINTESTINAL ------------------ */
  // Canine Periodontal Disease | Mild
    function generateDog1PeriodontalDiseaseTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Periodontal disease: Your dog has early dental disease. While brushing ${p.his} teeth is the best way to keep them clean, they will not remove the tartar & calculus that is already there. You can schedule a dental cleaning to completely remove calculus and then try brushing the teeth.`,
    `Until then, use a small dog toothbrush, medium/large dog toothbrush & animal safe toothpaste such as C.E.T. Start by having ${p.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then apply the pet safe toothpaste & let ${p.him} lick it off every day for a week. Finally, gently brush ${p.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
    `If your dog resists having ${p.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${p.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["MILD_PERIODONTAL_DISEASE"],
    boldKeys: [
    `PERIODONTAL_DISEASE_HEADER`,
    ],

    boldUnderlineKeys: [
    `DENTAL_BRUSHING`,
    `MILD_DENTAL_DZ`,
    `XYLITOL`,
    ],

    linkKeys: [
    `SMALL_DOG_TOOTHBRUSH_LINK`,
    `LARGE_TOOTHBRUSH_LINK`,
    `TOOTHPASTE_LINK`,
    `VOHC_DOG_LINK`,
    ],
    };
    }

  // Canine Periodontal Disease | Moderate
    function generateDog2PeriodontalDiseaseTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Periodontal disease: Your dog needs a Complete Oral Health Assessment and Treatment (COHAT) procedure. While brushing ${p.his} teeth is the best way to keep them clean, they will not remove the tartar & calculus that is already there. Schedule a dental cleaning within the next three months.`,
    `Brushing can still be performed right now but will be most effective after the next cleaning. Wait 3 weeks after ${p.his} next teeth cleaning before brushing the teeth to allow time for the mouth’s soreness to abate. Use a small dog toothbrush, canine toothbrush, or finger toothbrush & animal safe toothpaste such as C.E.T. Start by having ${p.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then apply the pet safe toothpaste & let ${p.him} lick it off every day for a week. Finally, gently brush ${p.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
    `If your dog resists having ${p.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${p.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["MODERATE_PERIODONTAL_DISEASE"],
    boldKeys: [
    `PERIODONTAL_DISEASE_HEADER`,
    ],

    boldUnderlineKeys: [
    `BRUSHING_LESS_EFFECTIVE`,
    `COHAT_RECOMMENDED`,
    `DENTAL_BRUSHING`,
    `SCHEDULE_DENTAL`,
    `XYLITOL`,
    ],

    linkKeys: [
    `SMALL_DOG_TOOTHBRUSH_LINK`,
    `LARGE_TOOTHBRUSH_LINK`,
    `TOOTHPASTE_LINK`,
    `VOHC_DOG_LINK`,
    ],
    };
    }

  // Canine Periodontal Disease | Severe
    function generateDog3PeriodontalDiseaseTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Periodontal disease: Your dog needs a Complete Oral Health Assessment and Treatment (COHAT) procedure. Brushing ${p.his} teeth is the best way to keep them clean, but it will not remove the tartar & calculus that is already there. In fact, brushing right now is not advised as it will likely cause ${p.him} pain and possibly bleeding given how severe the disease is. Schedule a dental cleaning within the next three months.`,
    `In the meantime, feed soft food such as wet food or dry food soaked in a few tablespoons of warm water 30 seconds prior to feeding. This will make it easier for your dog to chew. Wait 3 weeks after ${p.his} next teeth cleaning before brushing the teeth to allow time for the mouth’s soreness to abate. Use a small dog toothbrush, medium/large dog toothbrush, or finger toothbrush & animal safe toothpaste such as C.E.T. Start by having ${p.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then apply the pet safe toothpaste & let ${p.him} lick it off every day for a week. Finally, gently brush ${p.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
    `If your dog resists having ${p.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${p.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["SEVERE_PERIODONTAL_DISEASE"],
    boldKeys: [
    `PERIODONTAL_DISEASE_HEADER`,
    ],

    boldUnderlineKeys: [
    `COHAT_RECOMMENDED`,
    `DENTAL_BRUSHING`,
    `SCHEDULE_DENTAL`,
    `SOFTEN_FOOD`,
    `XYLITOL`,
    ],

    linkKeys: [
    `SMALL_DOG_TOOTHBRUSH_LINK`,
    `LARGE_TOOTHBRUSH_LINK`,
    `TOOTHPASTE_LINK`,
    `VOHC_DOG_LINK`,
    ],
    };
    }

  // Canine Periodontal Disease | Age Restricted
    function generateDog4PeriodontalDiseaseAgeTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Periodontal disease: Your dog shows signs of dental disease. However, older patients are more at risk of anesthesia complications. A routine dental cleaning is not advised in your dog for that reason unless it is performed with a dental specialist. The recommended clinic is Veterinary Dental Specialists. A referral to those clinics can be facilitated at your request.`,
    `In the meantime, feed soft food such as wet food or dry food soaked in a few tablespoons of warm water 30 seconds prior to feeding. This will make it easier for your dog to chew the food. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["PERIODONTAL_DISEASE"],
    boldKeys: [
    `PERIODONTAL_DISEASE_HEADER`,
    ],

    boldUnderlineKeys: [
    `SOFTEN_FOOD`,
    ],

    linkKeys: [
    `VOHC_DOG_LINK`,
    ],
    };
    }

  // Canine Periodontal Disease | Concurrent Disease
    function generateDog4PeriodontalDiseaseConcurrentDiseaseTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Periodontal disease: Your dog shows signs of dental disease. However, a dental cleaning is not recommended at this time until the current problem is dealt with. Once the other problem has been addressed, a dental cleaning can be performed. In the meantime, feed soft food such as wet food or dry food soaked in a few tablespoons of warm water 30 seconds prior to feeding. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["PERIODONTAL_DISEASE"],
    boldKeys: [
    `PERIODONTAL_DISEASE_HEADER`,
    ],

    boldUnderlineKeys: [
    `SOFTEN_FOOD`,
    ],

    linkKeys: [
    `VOHC_DOG_LINK`,
    ],
    };
    }

  // Canine Periodontal Disease | Heart Murmur
    function generateDog4PeriodontalDiseaseHeartMurmurTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Periodontal disease: Your dog shows signs of dental disease. However, a dental cleaning is not recommended at this time until the heart is further investigated. Once the other problem has been addressed, a dental cleaning can be performed. In the meantime, feed soft food such as wet food or dry food soaked in a few tablespoons of warm water 30 seconds prior to feeding. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["PERIODONTAL_DISEASE"],
    boldKeys: [
    `PERIODONTAL_DISEASE_HEADER`,
    ],

    boldUnderlineKeys: [
    `SOFTEN_FOOD`,
    ],

    linkKeys: [
    `VOHC_DOG_LINK`,
    ],
    };
    }

/* ------------------ CANINE DERMATOLOGY ------------------ */
  // Canine Atopic Dermatitis | Antihistamines 1
    function generateDogAtopicDermatitisMild1Template(sex) {
    const p = getPronoun(sex);

    const text = [
    `Atopic dermatitis: Unlike humans where allergies present in the respiratory tract (runny nose, sneezing/coughing, etc.), allergies in pets usually appear in the skin (shaking the head, chewing/licking the paws, scratching excessively, etc.). In fact, one of the most common causes of chronic ear infections is allergies.`,
    `At this time we will not be starting with daily oral medicine (Apoquel and Zenrelia) or monthly injectable medicine (Cytopoint). You can give over the counter antihistamines such as Benadryl 25mg (give up to 1 tablet per 25 lbs every 12 hours) or Zyrtec 10mg (give up to 1 tablet per 10 lbs every 12 - 24 hours). Potential side effects (drowsiness, increased drinking) are more common with Benadryl than Zyrtec.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER'
    ],

    boldUnderlineKeys: [
    `ALLERGY_EAR_RELATIONSHIP`,
    'ANTIHISTAMINE_DOSAGE1',
    ],
    };
    }

  // Canine Atopic Dermatitis | Antihistamines 2
    function generateDogAtopicDermatitisMild2Template(sex) {
    const p = getPronoun(sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies which you currently give over the counter antihistamines for. Continue to give Benadryl 25mg (give up to 1 tablet per 25 lbs every 12 hours) or Zyrtec 10mg (give up to 1 tablet per 10 lbs every 12 - 24 hours) as needed. If you feel like allergies are not well controlled, prescription medicine such as Cytopoint (an injection given every 4 - 8 weeks) or Apoquel OR Zenrelia (oral pills given every 24 hours) can be given for better control.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER',
    ],

    boldUnderlineKeys: [
    'ANTIHISTAMINE_DOSAGE2',
    ],
    };
    }

  // Canine Atopic Dermatitis | Apoquel 1
    function generateDogAtopicDermatitis1ApoquelTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Atopic dermatitis: Unlike humans where allergies presents in the respiratory tract (runny nose, sneezing, etc.), allergies in pets usually appears in the skin (shaking the head, chewing/licking the paws, scratching excessively, etc.). In fact, one of the most common causes of chronic ear infections is allergies. While antihistamines (Benadryl, Zyrtec, etc.) occasionally help, your dog shows signs of severe allergies.`,
    `Apoquel has been sent home to resolve allergies. Give as prescribed. If itching & scratching persists after two weeks, Cytopoint (an injection given every 4 - 8 weeks) can be tried instead. Alternatively, they can be given together to have a more powerful effect to control allergies. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`,
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER',
    ],

    boldUnderlineKeys: [
    `APOQUEL_STARTER`,
    `ANTIHISTAMINE_ADDITION`,
    ],

    greenKeys: [],
    linkKeys: [],
    };
    }

  // Canine Atopic Dermatitis | Apoquel 2, Maintenance
    function generateDogAtopicDermatitis2ApoquelTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies & gets Apoquel to control them. If itching & scratching persists, Cytopoint (an injection given every 4 - 8 weeks) can be tried instead. Alternatively, they can be given together to have a more powerful effect to control allergies. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER',
    ],

    boldUnderlineKeys: [
    `ANTIHISTAMINE_ADDITION`,
    ],
    };
    }

  // Canine Atopic Dermatitis | Apoquel 3, Add Cytopoint
    function generateDogAtopicDermatitis3ApoquelTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies & gets Apoquel to control them. However, Apoquel on its own doesn’t appear effective enough to control allergies. As such we will be adding Cytopoint to the plan. These medications improve the effectiveness of the other and are safe to give together.`,
    `Continue to give Apoquel as you’ve been doing. If you see full allergy control, you can try discontinuing Apoquel in 2 weeks to see if Cytopoint on its own can help control allergies. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER',
    ],

    boldUnderlineKeys: [
    `ANTIHISTAMINE_ADDITION`,
    `SYNERGISTIC_MEDICINE`,
    ],
    };
    }

  // Canine Atopic Dermatitis | Apoquel 4 Switch to Zenrelia
    function generateDogAtopicDermatitis4ApoquelTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies & gets Apoquel to control them. However, Apoquel doesn’t appear to be effective enough. We will be switching your dog to Zenrelia instead to see if this better controls allergies. Give daily for 1 month for best results. Do not give Zenrelia in the same 24 hours as Apoquel.`,
    `If itching & scratching persists, Cytopoint (an injection given every 4 - 8 weeks) can be tried instead. Alternatively, they can be given together to have a more powerful effect to control allergies. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER',
    ],

    boldUnderlineKeys: [
    `ANTIHISTAMINE_ADDITION`,
    `ZENRELIA_AND_APOQUEL_WARNING`,
    ],
    };
    }

  // Canine Atopic Dermatitis | Cytopoint 1
    function generateDogAtopicDermatitis1CytopointTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Atopic dermatitis: Unlike humans where allergies presents in the respiratory tract (runny nose, sneezing, etc.), allergies in pets usually appears in the skin (shaking the head, chewing/licking the paws, scratching excessively. etc.). In fact, one of the most common causes of chronic ear infections is allergies. While antihistamines (Benadryl, Zyrtec, etc.) occasionally help, your dog shows signs of severe allergies.`,
    `Cytopoint has been given in clinic to resolve allergies. It typically lasts 4 - 8 weeks. If itching & scratching occurs before 4 weeks, Apoquel OR Zenrelia (oral pills given once a day) can be sent home in addition to monthly Cytopoint injections to better control allergies.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER',
    ],

    boldUnderlineKeys: [
    `CYTOPOINT_STARTER`,
    ],
    };
    }

  // Canine Atopic Dermatitis | Cytopoint 2
    function generateDogAtopicDermatitis2CytopointTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies & gets Cytopoint injections to control them. The injection was given today & typically lasts 4 - 8 weeks. If the allergies return before 4 weeks, your dog may need Apoquel or Zenrelia in addition to Cytopoint. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER',
    ],

    boldUnderlineKeys: [
    `CYTOPOINT_ADDITIONAL_SUPPORT`,
    ],
    };
    }

  // Canine Atopic Dermatitis | Meds Declined
    function generateDogAtopicDermatitisMedsDeclinedTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Atopic dermatitis: Unlike humans where allergies presents in the respiratory tract (runny nose, sneezing, etc.), allergies in pets usually appears in the skin (shaking the head, chewing/licking the paws, scratching excessively. etc.). In fact, one of the most common causes of chronic ear infections is allergies.`,
    `Cytopoint (an injection given every 4 - 8 weeks) or either Apoquel OR Zenrelia (oral pills given every 24 hours) are more effective than over the counter medicine and are advised, but you have elected to try antihistamines first. You can give over the counter antihistamines such as Benadryl 25mg (give up to 1 tablet per 25 lbs every 12 hours) or Zyrtec 10mg (give up to 1 tablet per 10 lbs every 12 - 24 hours). Side effects (drowsiness, increased drinking) are more common with Benadryl than Zyrtec.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER',
    ],

    boldUnderlineKeys: [
    `ANTIHISTAMINE_DOSAGE1`,
    ],
    };
    }

  // Canine Atopic Dermatitis | Zenrelia 1
    function generateDogAtopicDermatitis1ZenreliaTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Atopic dermatitis: Unlike humans where allergies presents in the respiratory tract (runny nose, sneezing, etc.), allergies in pets usually appears in the skin (shaking the head, chewing/licking the paws, scratching excessively, etc.). In fact, one of the most common causes of chronic ear infections is allergies. While antihistamines (Benadryl, Zyrtec, etc.) occasionally help, your dog shows signs of severe allergies.`,
    `Zenrelia has been sent home to resolve allergies. Give as prescribed. If itching & scratching persists after two weeks, Cytopoint (an injection given every 4 - 8 weeks) can be tried instead. Alternatively, they can be given together to have a more powerful effect to control allergies.`,
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER',
    ],

    boldUnderlineKeys: [
    `ZENRELIA_STARTER`,
    ],

    greenKeys: [],
    linkKeys: [],
    };
    }

  // Canine Atopic Dermatitis | Zenrelia 2, Maintenance
    function generateDogAtopicDermatitis2ZenreliaTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies & gets Zenrelia to control them. If itching & scratching persists, Cytopoint (an injection given every 4 - 8 weeks) can be tried instead. Alternatively, they can be given together to have a more powerful effect to control allergies. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`,
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER',
    ],

    boldUnderlineKeys: [
    `ANTIHISTAMINE_ADDITION`,
    ],
    };
    }

  // Canine Atopic Dermatitis | Zenrelia 3, Add Cytopoint
    function generateDogAtopicDermatitis3ZenreliaTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies & gets Zenrelia to control them. However, Zenrelia on its own doesn’t appear effective enough to control allergies. As such we will be adding Cytopoint to the plan. These medications improve the effectiveness of the other and are safe to give together.`,
    `Continue to give Zenrelia as you’ve been doing. If you see full allergy control, you can try discontinuing Zenrelia in 2 weeks to see if Cytopoint on its own can help control allergies. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER',
    ],

    boldUnderlineKeys: [
    `ANTIHISTAMINE_ADDITION`,
    `SYNERGISTIC_MEDICINE`,
    ],
    };
    }

/* ------------------ CANINE MUSCULOSKELETAL ------------------ */
  // Canine Osteoarthritis | 1st NSAID, Initial
    function generateDogOsteoarthritis1NSAIDTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Osteoarthritis: Arthritis was detected in your dog’s joints. The most common sign of this is being slow after waking up/laying down for a while or being sore after walks. There are three options for treatment: monthly injectable medicine, twice daily oral pain medicine, and joint supplements. You’ve elected to try a non-steroidal anti-inflammatory drug (NSAID) to reduce pain & inflammation from arthritis. Bloodwork is required every 6 - 12 months while on NSAIDs.`,
    `Alternatives include Librela (an injection given in clinic once a month), gabapentin (oral capsules), & joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size). Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["OSTEOARTHRITIS"],
    boldKeys: [
    `OSTEOARTHRITIS_HEADER`,
    ],

    boldUnderlineKeys: [
    `ARTHRITIS_DETECTED`,
    `NSAID_LABWORK_REQUIREMENT`,
    ],
    };
    }

  // Canine Osteoarthritis | 2nd NSAID, Maintenance
    function generateDogOsteoarthritis2NSAIDTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Osteoarthritis: Your dog is known to have arthritis & is currently on a non-steroidal anti-inflammatory drug (NSAID) to reduce pain & inflammation from arthritis. Bloodwork is required every 6 - 12 months while on NSAIDs. Alternatives include Librela (an injection given in clinic once a month), gabapentin (oral capsules), & joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size). Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["OSTEOARTHRITIS"],
    boldKeys: [
    `OSTEOARTHRITIS_HEADER`,
    ],

    boldUnderlineKeys: [
    `NSAID_MONITORING`,
    ],

    linkKeys: [

    ],
    };
    }

  // Canine Osteoarthritis | 3rd NSAID, Switch NSAIDs
    function generateDogOsteoarthritis3NSAIDTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Osteoarthritis: Your dog is known to have arthritis & is currently on a non-steroidal anti-inflammatory drug (NSAID) to reduce pain & inflammation from arthritis. We will be switching to a different NSAID to see if better control is provided while still being safe for your pet. Bloodwork is required every 6 - 12 months while on NSAIDs.`,
    `If we still don’t see the improvement we’d like, we can discuss adding on alternatives such as Librela (an injection given in clinic once a month), gabapentin (oral capsules), & joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size). Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["OSTEOARTHRITIS"],
    boldKeys: [
    `OSTEOARTHRITIS_HEADER`,
    ],

    boldUnderlineKeys: [
    `NSAID_MONITORING`,
    ],

    linkKeys: [

    ],
    };
    }

  // Canine Osteoarthritis | 1st Gabapentin
    function generateDogOsteoarthritis1GabapentinTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Osteoarthritis: Arthritis was detected in your dog’s joints. The most common sign of this is being slow after waking up/laying down for a while or being sore after walks. There are three options for treatment: monthly injectable medicine, twice daily oral pain medicine, and joint supplements.`,
    `Librela is an injection given once a month that controls arthritis in most patients. However, it may take 2 - 3 months before improvement is seen. Instead, immediate relief can be provided via NSAIDs such as carprofen & grapiprant. Bloodwork is recommended every 6 - 12 months while on NSAIDs. `,
    `You’ve elected to try gabapentin. While less effective than NSAIDs, they do not require bloodwork and still offer excellent pain control. Joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size) are a more natural alternative that you can add on to improve mobility, but they do not reduce pain on their own. Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["OSTEOARTHRITIS"],
    boldKeys: [
    'OSTEOARTHRITIS_HEADER'
    ],

    boldUnderlineKeys: [ 
    `ARTHRITIS_DETECTED`, 
    `GABAPENTIN_DECISION`,
    ],
    };
    }

  // Canine Osteoarthritis | 2nd Gabapentin, Continue
    function generateDogOsteoarthritis2GabapentinTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Osteoarthritis: Your dog is known to have arthritis & is currently on gabapentin. Alternatives to gabapentin include Librela, an injection that can be given once a month with minimal side effects. However, it may take 2 - 3 months before improvement is seen. Non-steroidal anti-inflammatory drugs such as carprofen & grapiprant can also control arthritis so long as steroids aren’t currently being given. Bloodwork is recommended every 6 - 12 months while on NSAIDs.`,
    `Joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size) can also be added to any treatment plan to increase mobility but will not control pain. Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["OSTEOARTHRITIS"],
    boldKeys: [ 
    `OSTEOARTHRITIS_HEADER`,
    ],

    boldUnderlineKeys: [ 
    `ARTHRITIS_WEIGHT_MANAGEMENT`,
    `GABAPENTIN_ARTHRITIS_MANAGEMENT`,
    ],
    };
    }

  // Canine Osteoarthritis | 1st Joint Supplements
    function generateDogOsteoarthritis1JointSupplementsTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Osteoarthritis: Arthritis was detected in your dog’s joints. The most common sign of this is being slow after waking up/laying down for a while or being sore after walks. There are three options for treatment: monthly injectable medicine, twice daily oral pain medicine, and joint supplements. Immediate relief can be provided via gabapentin or non-steroidal anti-inflammatory drugs (NSAIDs) such as carprofen & grapiprant. NSAIDs are more effective at controlling arthritis than gabapentin & require bloodwork every 6 months. Librela is an injection given once a month that controls arthritis in most patients. However, it may take 2 - 3 months before improvement is seen.`,
    `At this time you’ve elected to try joint supplements. Joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size) are a more natural alternative to increase mobility but do not reduce pain. Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["OSTEOARTHRITIS"],
    boldKeys: [ 
    "OSTEOARTHRITIS_HEADER" 
    ],

    boldUnderlineKeys: [ 
    "ARTHRITIS_DETECTED",
    "ARTHRITIS_WEIGHT_MANAGEMENT",
    "JOINT_SUPPLEMENTS_DECISION",
    ],
    };
    }

  // Canine Osteoarthritis | 2nd Joint Supplements, Continue
    function generateDogOsteoarthritis2JointSupplementsTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Osteoarthritis: Your dog is known to have arthritis & currently gets joint supplements. Joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size) are a more natural alternative to increase mobility but do not reduce pain. Non-steroidal anti-inflammatory drugs (NSAIDs) such as carprofen & grapiprant provide immediate relief from arthritis so long as steroids aren’t currently being given. Bloodwork is required every 6 months while on NSAIDs. `,
    `Gabapentin can be used in addition to or instead of NSAIDs with minimal side effects, though it isn’t as effective as NSAIDs. Finally, the Librela injection can be given once a month though it can take 2 - 3 months to see improvement. If you have concerns about your dog’s arthritis, contact the clinic & we can discuss which medicine is best for your dog. Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["OSTEOARTHRITIS"],
    boldKeys: [ 
    "OSTEOARTHRITIS_HEADER" 
    ],

    boldUnderlineKeys: [ 
    "ARTHRITIS_WEIGHT_MANAGEMENT",
    "JOINT_SUPPLEMENTS_KNOWN",
    ],
    };
    }

  // Canine Osteoarthritis | 1st Librela
    function generateDogOsteoarthritis1LibrelaTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Osteoarthritis: Arthritis was detected in your dog’s joints. The most common sign of this is being slow after waking up/laying down for a while or being sore after walks. There are three options for treatment: monthly injectable medicine, twice daily oral pain medicine, and joint supplements.`,
    `Librela, an injection given once a month that controls arthritis, was given today. Watch for signs of reaction including vomiting, diarrhea, lethargy, & excessive panting/fever. If signs are seen, bring your dog back immediately as these are signs of a reaction. Librela may take 2 - 3 months before full effects are seen, so oral medicine (gabapentin, carprofen, grapiprant, etc.) can be used in the meantime if necessary.`,
    `You can continue giving joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size) to help improve joint mobility. Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["OSTEOARTHRITIS"],
    boldKeys: [ 
    "OSTEOARTHRITIS_HEADER" 
    ],

    boldUnderlineKeys: [ 
    "ARTHRITIS_DETECTED",
    "ARTHRITIS_WEIGHT_MANAGEMENT" ,
    "LIBRELA_ADVERSE_RXN",
    ],
    };
    }

  // Canine Osteoarthritis | 2nd Librela
    function generateDogOsteoarthritis2LibrelaTemplate(sex) {
    const p = getPronoun(sex);
    const text = [
    `Osteoarthritis: Your dog is known to have arthritis & currently gets Librela injections. Watch for signs of reaction including vomiting, diarrhea, lethargy, & excessive panting/fever. Bring your dog back immediately if any are seen. Joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size) can be added in addition to further improve mobility. Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    text,
    diagnoses: ["OSTEOARTHRITIS"],
    boldKeys: [ 
    "OSTEOARTHRITIS_HEADER" 
    ],

    boldUnderlineKeys: [ 
    "ARTHRITIS_WEIGHT_MANAGEMENT",
    "LIBRELA_ADVERSE_RXN",
    ],
    };
    }  

/* ------------------ TEMPLATE DEFINITIONS ------------------ */
  // Main Function
    const RAW_TEMPLATE_DEFINITIONS = {
    '/ieolibrary': function () {
    return {
    text: buildKeywordLibrary(),
    boldKeys: [],
    boldUnderlineKeys: [],
    greenKeys: [],
    linkKeys: []
    };
    },

    '/generatemedicinetable': () => ({
    text: "",
    customAction: generateMedicineTableFromBuffer
    }),

  // Puppy Wellness
    '/cReset': () => generateCanineResetTemplate(),
    '/fReset': () => generateFelineResetTemplate(),
    '/c8wksSmallMale': () => generate8WkWellnessTemplate('small', 'male'),
    '/c8wksSmallMales': () => generate8WkWellnessTemplate('small', 'male', 'plural'),
    '/c8wksLargeMale': () => generate8WkWellnessTemplate('large', 'male'),
    '/c8wksLargeMales': () => generate8WkWellnessTemplate('large', 'male', 'plural'),
    '/c8wksFemale': () => generate8WkWellnessTemplate('small', 'female'),
    '/c8wksFemales': () => generate8WkWellnessTemplate('small', 'female', 'plural'),

    '/c12wksSmallMale': () => generate12WkWellnessTemplate('small', 'male'),
    '/c12wksSmallMales': () => generate12WkWellnessTemplate('small', 'male', 'plural'),
    '/c12wksLargeMale': () => generate12WkWellnessTemplate('large', 'male'),
    '/c12wksLargeMales': () => generate12WkWellnessTemplate('large', 'male', 'plural'),
    '/c12wksFemale': () => generate12WkWellnessTemplate('small', 'female'),
    '/c12wksFemales': () => generate12WkWellnessTemplate('small', 'female', 'plural'),

    '/c16wksSmallMale': () => generate16WkWellnessTemplate('small', 'male'),
    '/c16wksSmallMales': () => generate16WkWellnessTemplate('small', 'male', 'plural'),
    '/c16wksLargeMale': () => generate16WkWellnessTemplate('large', 'male'),
    '/c16wksLargeMales': () => generate16WkWellnessTemplate('large', 'male', 'plural'),
    '/c16wksFemale': () => generate16WkWellnessTemplate('small', 'female'),
    '/c16wksFemales': () => generate16WkWellnessTemplate('small', 'female', 'plural'),

  // Canine Adult Wellness
    '/cInitialAdultMale': () => generateInitialAdultTemplate('male'),
    '/cInitialAdultMales': () => generateInitialAdultTemplate('male', 'plural'),
    '/cInitialAdultFemale': () => generateInitialAdultTemplate('female'),
    '/cInitialAdultFemales': () => generateInitialAdultTemplate('female', 'plural'),
    '/c1yearMale': () => generate1YearAdultTemplate('male'),
    '/c1yearMales': () => generate1YearAdultTemplate('male', 'plural'),
    '/c1yearFemale': () => generate1YearAdultTemplate('female'),
    '/c1yearFemales': () => generate1YearAdultTemplate('female', 'plural'),
    '/c2yearMale': () => generate2YearAdultTemplate('male'),
    '/c2yearMales': () => generate2YearAdultTemplate('male', 'plural'),
    '/c2yearFemale': () => generate2YearAdultTemplate('female'),
    '/c2yearFemales': () => generate2YearAdultTemplate('female', 'plural'),
    '/c2yearLeptoMale': () => generate2YearLeptoTemplate('male'),
    '/c2yearLeptoMales': () => generate2YearLeptoTemplate('male', 'plural'),
    '/c2yearLeptoFemale': () => generate2YearLeptoTemplate('female'),
    '/c2yearLeptoFemales': () => generate2YearLeptoTemplate('female', 'plural'),
    '/c7yearMale': () => generate7YearAdultTemplate('male'),
    '/c7yearMales': () => generate7YearAdultTemplate('male', 'plural'),
    '/c7yearFemale': () => generate7YearAdultTemplate('female'),
    '/c7yearFemales': () => generate7YearAdultTemplate('female', 'plural'),
    '/c7yearLeptoMale': () => generate7YearLeptoTemplate('male'),
    '/c7yearLeptoMales': () => generate7YearLeptoTemplate('male', 'plural'),
    '/c7yearLeptoFemale': () => generate7YearLeptoTemplate('female'),
    '/c7yearLeptoFemales': () => generate7YearLeptoTemplate('female', 'plural'),

    '/cOverweightMale': () => generateCanineOverweightTemplate('male'),
    '/cOverweightMales': () => generateCanineOverweightTemplate('male', 'plural'),
    '/cOverweightFemale': () => generateCanineOverweightTemplate('female'),
    '/cOverweightFemales': () => generateCanineOverweightTemplate('female', 'plural'),
    '/cOverweight2Male': () => generateCanineOverweight2Template('male'),
    '/cOverweight2Males': () => generateCanineOverweight2Template('male', 'plural'),
    '/cOverweight2Female': () => generateCanineOverweight2Template('female'),
    '/cOverweight2Females': () => generateCanineOverweight2Template('female', 'plural'),
    '/cHealthyWeightMale': () => generateDogHealthyWeightTemplate('male'),
    '/cHealthyWeightMales': () => generateDogHealthyWeightTemplate('male', 'plural'),
    '/cHealthyWeightFemale': () => generateDogHealthyWeightTemplate('female'),
    '/cHealthyWeightFemales': () => generateDogHealthyWeightTemplate('female', 'plural'),
    '/cUnderweightMale': () => generateDogUnderweightTemplate('male'),
    '/cUnderweightMales': () => generateDogUnderweightTemplate('male', 'plural'),
    '/cUnderweightFemale': () => generateDogUnderweightTemplate('female'),
    '/cUnderweightFemales': () => generateDogUnderweightTemplate('female', 'plural'),

  // Canine Ophthalmology
    '/cBlind0Partial': () => generateDogBlind0PartialTemplate(),
    '/cBlind1': () => generateDogBlind1Template(),
    '/cBlind2Known': () => generateDogBlind2KnownTemplate(),
    '/cCherryEye': () => generateDogCherryEyeTemplate(),
    '/cCherryEyes': () => generateDogCherryEyeTemplate('male','plural'),

  // Canine Cardiology
    '/cHeartMurmur0': () => generateDogHeartMurmur0Template(),
    '/cHeartMurmur1RadiographsNormal': () => generateDogHeartMurmur1RadiographsNormalTemplate(),
    '/cHeartMurmur1Cardiomegaly': () => generateDogHeartMurmur1CardiomegalyTemplate(),
    '/cHeartMurmur3Known': () => generateDogHeartMurmur3KnownTemplate(),

  // Gastrointestinal
    '/cPeriodontalDisease1Male': () => generateDog1PeriodontalDiseaseTemplate('male'),
    '/cPeriodontalDisease1Female': () => generateDog1PeriodontalDiseaseTemplate('female'),
    '/cPeriodontalDisease2Male': () => generateDog2PeriodontalDiseaseTemplate('male'),
    '/cPeriodontalDisease2Female': () => generateDog2PeriodontalDiseaseTemplate('female'),
    '/cPeriodontalDisease3Male': () => generateDog3PeriodontalDiseaseTemplate('male'),
    '/cPeriodontalDisease3Female': () => generateDog3PeriodontalDiseaseTemplate('female'),
    '/cPeriodontalDisease4AgeMale': () => generateDog4PeriodontalDiseaseAgeTemplate('male'),
    '/cPeriodontalDisease4AgeFemale': () => generateDog4PeriodontalDiseaseAgeTemplate('female'),
    '/cPeriodontalDisease4ConcurrentDiseaseMale': () => generateDog4PeriodontalDiseaseConcurrentDiseaseTemplate('male'),
    '/cPeriodontalDisease4ConcurrentDiseaseFemale': () => generateDog4PeriodontalDiseaseConcurrentDiseaseTemplate('female'),
    '/cPeriodontalDisease4HeartMurmurMale': () => generateDog4PeriodontalDiseaseHeartMurmurTemplate('male'),
    '/cPeriodontalDisease4HeartMurmurFemale': () => generateDog4PeriodontalDiseaseHeartMurmurTemplate('female'),

  // Musculoskeletal
    '/cOsteoarthritis1NSAID': () => generateDogOsteoarthritis1NSAIDTemplate(),
    '/cOsteoarthritis2NSAID': () => generateDogOsteoarthritis2NSAIDTemplate(),
    '/cOsteoarthritis3NSAID': () => generateDogOsteoarthritis3NSAIDTemplate(),
    '/cOsteoarthritis1Gabapentin': () => generateDogOsteoarthritis1GabapentinTemplate(),
    '/cOsteoarthritis2Gabapentin': () => generateDogOsteoarthritis2GabapentinTemplate(),
    '/cOsteoarthritis1JointSupplements': () => generateDogOsteoarthritis1JointSupplementsTemplate(),
    '/cOsteoarthritis2JointSupplements': () => generateDogOsteoarthritis2JointSupplementsTemplate(),
    '/cOsteoarthritis1Librela': () => generateDogOsteoarthritis1LibrelaTemplate(),
    '/cOsteoarthritis2Librela': () => generateDogOsteoarthritis2LibrelaTemplate(),

  // Immunology
    '/cVaccineInformation': () => generateDogVaccineInformationTemplate(),

  // Dermatology/
    '/cAtopicDermatitis1Antihistamines': () => generateDogAtopicDermatitisMild1Template(),
    '/cAtopicDermatitis2Antihistamines': () => generateDogAtopicDermatitisMild2Template(),
    '/cAtopicDermatitis1Apoquel': () => generateDogAtopicDermatitis1ApoquelTemplate(),
    '/cAtopicDermatitis2Apoquel': () => generateDogAtopicDermatitis2ApoquelTemplate(),
    '/cAtopicDermatitis3Apoquel': () => generateDogAtopicDermatitis3ApoquelTemplate(),
    '/cAtopicDermatitis4Apoquel': () => generateDogAtopicDermatitis4ApoquelTemplate(),
    '/cAtopicDermatitis1Cytopoint': () => generateDogAtopicDermatitis1CytopointTemplate(),
    '/cAtopicDermatitis2Cytopoint': () => generateDogAtopicDermatitis2CytopointTemplate(),
    '/cAtopicDermatitisMedsDeclined': () => generateDogAtopicDermatitisMedsDeclinedTemplate(),
    '/cAtopicDermatitis1Zenrelia': () => generateDogAtopicDermatitis1ZenreliaTemplate(),
    '/cAtopicDermatitis2Zenrelia': () => generateDogAtopicDermatitis2ZenreliaTemplate(),
    '/cAtopicDermatitis3Zenrelia': () => generateDogAtopicDermatitis3ZenreliaTemplate(),
    };

  // Template Definitions
    const TEMPLATE_DEFINITIONS = {};

    Object.keys(RAW_TEMPLATE_DEFINITIONS).forEach(key => {
    TEMPLATE_DEFINITIONS[key.toLowerCase()] = RAW_TEMPLATE_DEFINITIONS[key];
    });

    function resolveTemplate(keyword) {
    const normalizedKeyword = keyword.toLowerCase();
    const templateFn = TEMPLATE_DEFINITIONS[normalizedKeyword];

    if (!templateFn) {
    throw new Error('Template not found: ' + keyword);
    }

    return templateFn();
    }

    function buildKeywordLibrary() {
    const keys = Object.keys(TEMPLATE_DEFINITIONS)
    .filter(k => k !== '/ieolibrary');

    let output = [];
    output.push("Dr. I.E. Osadiaye's Keyword Library");
    output.push('--------------------');

    keys.forEach(key => {
    output.push(key);
    });

    return output.join('\n');
    }

/* ------------------ EXPAND KEYWORDS ------------------ */
  // Main Function 
    function expandKeywords() {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    // Improved pattern to capture brackets for medications
    const combinedPattern = '\\/[a-zA-Z0-9]+(\\[.*?\\])*'; 

    let searchResult = null;
    const matches = [];

    // 1. COLLECT ALL KEYWORDS
    while ((searchResult = body.findText(combinedPattern, searchResult))) {
    const textElement = searchResult.getElement().asText();
    const matchedText = textElement.getText().substring(
    searchResult.getStartOffset(),
    searchResult.getEndOffsetInclusive() + 1
    ).trim();

    matches.push({
    text: matchedText,
    normalized: matchedText.toLowerCase(),
    element: textElement,
    start: searchResult.getStartOffset(),
    end: searchResult.getEndOffsetInclusive()
    });
    }

    // 2. DELETE KEYWORDS FROM DOC (Backwards to maintain indices)
    for (let i = matches.length - 1; i >= 0; i--) {
    const m = matches[i];
    const parent = m.element.getParent();

    // Delete the actual keyword text (e.g., /cReset)
    m.element.deleteText(m.start, m.end);

    // Check if the paragraph is now totally empty
    if (parent.asParagraph().getText().trim() === "") {
    try {
    // Only attempt removal if there is more than one element in the doc
    if (body.getNumChildren() > 1) {
    body.removeChild(parent);
    }
    } catch (e) {
    // If Google still refuses to delete it (e.g., it's the last line),
    // we just "catch" the error here and let the script keep running.
    console.warn("Skipped paragraph removal to prevent crash: " + e.message);
    }
    }
    }

    // 3. PRIORITY PASS: Handle Resets First
    const resetMatch = matches.find(m => m.normalized.endsWith('reset'));
    if (resetMatch) {
    const templateFn = TEMPLATE_DEFINITIONS[resetMatch.normalized];
    if (templateFn) {
    const template = templateFn();

    // CLEAR the body for a true reset
    body.clear(); 
    // Insert the structural template at the top (index 0)
    insertTemplateAtIndex(body, template, 0);

    // Run any custom logic if it exists
    if (template.customAction) template.customAction();
    }
    }

    // 4. SECOND PASS: Process medications and buffer standard templates
    matches.forEach(m => {
    if (m.normalized.endsWith('reset')) return; // Already handled

    const templateFn = TEMPLATE_DEFINITIONS[m.normalized];
    if (templateFn) {
    const template = templateFn();
    let rank = 99;
    if (template.diagnoses && template.diagnoses.length > 0) {
    bufferDiagnoses(template.diagnoses);
    const firstKey = template.diagnoses[0];
    rank = (DIAGNOSIS_REGISTRY[firstKey] && DIAGNOSIS_REGISTRY[firstKey].rank) || 99;
    }
    bufferTemplate(template, rank);
    if (template.customAction) template.customAction();
    }

    // Medication Command Logic
    if (m.normalized.startsWith("/c") && !TEMPLATE_DEFINITIONS[m.normalized]) {
    const medRow = processMedicationCommand(m.text);
    if (medRow) TABLE_ROW_BUFFER.push(medRow);
    }
    });

    // 5. FINAL FLUSH: Insert into the now-existing structure
    insertDiagnosesIntoDocument();
    insertTemplatesIntoDocument();
    if (TABLE_ROW_BUFFER.length > 0) generateMedicineTableFromBuffer();
    }

  // Regex Helpers
    function escapeForRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

  // Insert Template at Index
    function insertTemplateAtIndex(body, template, insertIndex) {
    const paragraphs = template.text.split('\n');
    const insertedParagraphs = [];
    const g = getGrammar('wellness', template.plurality || 'singular', template.sex || 'male');

    //  Helper to resolve Registry Keys ---
    const resolve = (keys) => (keys || []).map(key => {
    const entry = FORMAT_REGISTRY[key];
    if (!entry) return null;
    return typeof entry === 'function' ? entry(g || { he:'he', him:'him', his:'his' }) : entry;
    }).filter(Boolean);

    const resolvedBoldOnly = resolve(template.boldKeys);
    const resolvedBoldUnderline = resolve(template.boldUnderlineKeys);
    const resolvedItalic = resolve(template.italicKeys);
    const resolvedGreen = resolve(template.greenKeys);
    const resolvedRed = resolve(template.redKeys);
    const resolvedLinks = resolve(template.linkKeys);
    const resolvedTitle = resolve(template.titleKeys);
    const resolvedDoubleSpaced = resolve(template.doubleSpacedKeys);

    //  1. Insert Paragraphs & Set Paragraph-Level Styles ---
    paragraphs.forEach((paraText, i) => {
    const p = body.insertParagraph(insertIndex + i, paraText);

    // Default alignment and indentation
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    p.setIndentFirstLine(36);

    // Determine Line Spacing
    let spacing = template.blockLineSpacing || 2.0; // Default to double

    // Check if this specific line matches a Title Key
    if (resolvedTitle.some(t => paraText.includes(t))) {
    p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    spacing = STYLE_REGISTRY.title.lineSpacing; // 2.0
    }

    // Check if this specific line matches a Double Spaced Key (Review Request)
    if (resolvedDoubleSpaced.some(ds => paraText.includes(ds))) {
    spacing = STYLE_REGISTRY.doubleSpaced.lineSpacing; // 2.0
    }

    p.setLineSpacing(spacing);
    insertedParagraphs.push(p);
    });

    //  2. Apply Text-Level Formatting (Character Styles) ---
    insertedParagraphs.forEach(p => {
    const t = p.editAsText();
    const text = p.getText();
    if (!text.trim()) return;

    // Apply Title Size
    resolvedTitle.forEach(titleText => {
    if (text.includes(titleText)) {
    t.setFontSize(0, text.length - 1, 20);
    }
    });

    // Bold, Italic, Colors from Registry
    resolvedBoldOnly.forEach(str => applyFormattingOnce(t, str, STYLE_REGISTRY.bold));
    resolvedBoldUnderline.forEach(str => applyFormattingOnce(t, str, STYLE_REGISTRY.boldUnderline));
    resolvedItalic.forEach(str => applyFormattingOnce(t, str, STYLE_REGISTRY.italic));
    resolvedGreen.forEach(str => applyFormattingOnce(t, str, STYLE_REGISTRY.green));
    resolvedRed.forEach(str => applyFormattingOnce(t, str, STYLE_REGISTRY.red));

    // Hyperlinks
    resolvedLinks.forEach(link => {
    let startIndex = text.indexOf(link.text);
    while (startIndex !== -1) {
    t.setLinkUrl(startIndex, startIndex + link.text.length - 1, link.url);
    startIndex = text.indexOf(link.text, startIndex + link.text.length);
    }
    });
    });

    //  3. Buffer Table Data ---
    if (template.table) {
    // If table data is an array (multiple rows), spread it into the buffer
    template.table.data.forEach((row, idx) => {
    TABLE_ROW_BUFFER.push({
    rowData: row,
    color: template.table.colorRows ? template.table.colorRows[idx] : null
    });
    });
    }

    return insertedParagraphs;
    }

/* ------------------ TABLE INSERTION ------------------ */
  // Main Function
    function insertTableAtIndex(tableTemplate, insertIndex) {
    const body = DocumentApp.getActiveDocument().getBody();
    const table = body.insertTable(insertIndex);

    tableTemplate.data.forEach((rowData, rowIndex) => {
    const tableRow = table.appendTableRow();
    rowData.forEach((cellText, colIndex) => {
    const cell = tableRow.appendTableCell('');
    const baseParagraph = cell.getChild(0).asParagraph();
    baseParagraph.clear();

    if (tableTemplate.firstLineBoldUnderlineInstructions && colIndex === 1 && cellText.includes('\n')) {
    const parts = cellText.split('\n');

    const p1 = baseParagraph;
    p1.setText(parts[0]);
    const t1 = p1.editAsText();
    t1.setBold(true).setUnderline(true); // Bold + underline first line

    // Rest of the instructions normal
    for (let i = 1; i < parts.length; i++) {
    const p = cell.appendParagraph(parts[i]);
    const t = p.editAsText();
    t.setBold(false).setUnderline(false);
    }
    } else {
    // single line, normal formatting
    baseParagraph.setText(cellText);
    const t = baseParagraph.editAsText();
    t.setBold(false).setUnderline(false);
    }
    });

    if (tableTemplate.colorRows && tableTemplate.colorRows[rowIndex] !== undefined) {
    const color = tableTemplate.colorRows[rowIndex];
    for (let c = 0; c < tableRow.getNumCells(); c++) {
    tableRow.getCell(c).setBackgroundColor(color);
    }
    }
    });

    if (tableTemplate.boldTopRow) {
    const headerRow = table.getRow(0);
    for (let c = 0; c < headerRow.getNumCells(); c++) {
    const cell = headerRow.getCell(c);
    for (let pIndex = 0; pIndex < cell.getNumChildren(); pIndex++) {
    const paragraph = cell.getChild(pIndex).asParagraph();
    const t = paragraph.editAsText();
    t.setBold(true);
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);
    }
    }
    }

/* ------------------ TABLE BUFFER ------------------ */
  // Main Function
    function generateMedicineTableFromBuffer() {
    const body = DocumentApp.getActiveDocument().getBody();

    // 1. Exit if the buffer is totally empty
    if (!TABLE_ROW_BUFFER || !TABLE_ROW_BUFFER.length) {
    return; 
    }

    let mergedRows = [];
    let mergedColorRows = [];

    // 2. Safely extract data from the buffer
    TABLE_ROW_BUFFER.forEach((bufferedItem) => {
    // Check if rowData exists and is an array (the actual medication row)
    if (bufferedItem && Array.isArray(bufferedItem.rowData)) {
    mergedRows.push(bufferedItem.rowData);
    mergedColorRows.push(bufferedItem.color || null);
    }
    });

    // 3. Exit if no valid rows were found after processing
    if (mergedRows.length === 0) {
    TABLE_ROW_BUFFER = [];
    return;
    }

    // 4. Define the Header (Standard for all medication tables)
    const headerRow = ["Medication", "Instructions", "Class", "Side Effects"];
    const headerColor = null; // Transparent

    // 5. Deduplicate and Sort
    const seen = new Set();
    const combined = [];

    mergedRows.forEach((row, i) => {
    const key = row.join("|").toLowerCase().trim();
    if (!seen.has(key)) {
    seen.add(key);
    combined.push({
    row: row,
    color: mergedColorRows[i]
    });
    }
    });

    // Sort alphabetically by Medication Name (index 0)
    combined.sort((a, b) => a.row[0].toLowerCase().localeCompare(b.row[0].toLowerCase()));

    const sortedRows = combined.map(obj => obj.row);
    const sortedColors = combined.map(obj => obj.color);

    // 6. Final Table Template
    const mergedTableTemplate = {
    data: [headerRow, ...sortedRows],
    colorRows: [headerColor, ...sortedColors],
    boldTopRow: true,
    firstLineBoldUnderlineInstructions: true
    };

    // 7. Insert or Update in Document
    const existingTable = findExistingMedicineTable();

    if (existingTable) {
    appendRowsToExistingTable(existingTable, mergedTableTemplate);
    } else {
    // Insert at the very end of the document
    insertTableAtIndex(mergedTableTemplate, body.getNumChildren());
    }

    // 8. Clear the buffer for the next action
    TABLE_ROW_BUFFER = [];
    }

/* ------------------ Detect Medicine Table ------------------ */
  // Main Function
    function findExistingMedicineTable() {
    const body = DocumentApp.getActiveDocument().getBody();
    const tables = body.getTables();

    for (let table of tables) {
    if (table.getNumRows() === 0) continue;

    const firstRow = table.getRow(0);
    if (firstRow.getNumCells() < 4) continue;

    const headerText = [
    firstRow.getCell(0).getText(),
    firstRow.getCell(1).getText(),
    firstRow.getCell(2).getText(),
    firstRow.getCell(3).getText()
    ].join("|").toLowerCase();

    if (headerText.includes("medication") &&
    headerText.includes("instructions") &&
    headerText.includes("class") &&
    headerText.includes("side effects")) {
    return table;
    }
    }

    return null;
    }

/* ------------------ Append Medicine Table Rows ------------------ */
  // Main Function
    function appendRowsToExistingTable(table, tableTemplate) {

    const existingRows = new Set();

    for (let r = 1; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    const key = [
    row.getCell(0).getText(),
    row.getCell(1).getText(),
    row.getCell(2).getText(),
    row.getCell(3).getText()
    ].join("|").toLowerCase().trim();

    existingRows.add(key);
    }

    for (let i = 1; i < tableTemplate.data.length; i++) {

    const rowData = tableTemplate.data[i];
    const key = rowData.join("|").toLowerCase().trim();

    if (existingRows.has(key)) continue;

    const tableRow = table.appendTableRow();

    rowData.forEach((cellText, c) => {
    const cell = tableRow.appendTableCell(cellText);

    const paragraph = cell.getChild(0).asParagraph();
    const text = paragraph.editAsText();

    text.setBold(false);
    text.setUnderline(false);
    });

    const rowColor = tableTemplate.colorRows[i];

    if (rowColor) {
    for (let c = 0; c < tableRow.getNumCells(); c++) {
    tableRow.getCell(c).setBackgroundColor(rowColor);
    }
    }
    }

    sortMedicineTable(table);
    }

/* ------------------ SORT MEDICINE TABLE ------------------ */
  // Main Function
    function sortMedicineTable(table) {
    if (table.getNumRows() <= 2) return;

    const rows = [];

    // Collect rows with their current index
    for (let r = 1; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    const medName = row.getCell(0).getText().toLowerCase();
    rows.push({
    rowObj: row,   // store actual row object
    name: medName,
    index: r
    });
    }

    // Sort rows alphabetically
    rows.sort((a, b) => a.name.localeCompare(b.name));

    // Move rows into correct order
    rows.forEach((entry, newPosition) => {
    const desiredIndex = newPosition + 1; // header is row 0
    const currentRow = table.getRow(entry.index);

    if (entry.index !== desiredIndex) {
    const copy = currentRow.copy();
    table.insertTableRow(desiredIndex, copy);

    if (entry.index > desiredIndex) {
    table.removeRow(entry.index + 1);
    } else {
    table.removeRow(entry.index);
    }
    }
    });

    // ---  Apply first-line bold + underline to column 1 ---
    for (let r = 1; r < table.getNumRows(); r++) {
    const cell = table.getRow(r).getCell(1);
    const paragraph = cell.getChild(0).asParagraph();
    const text = paragraph.getText();
    const t = paragraph.editAsText();

    const firstLine = text.split("\n")[0];
    t.setBold(0, firstLine.length - 1, true);
    t.setUnderline(0, firstLine.length - 1, true);
    }
    }

/* ------------------ FORMATTING HELPERS ------------------ */
  // Main Function
    function applyFormattingOnce(textElement, match, style) {
    const fullText = textElement.getText();

    let regex;
    if (match instanceof RegExp) {
    // Preserve existing flags but force case-insensitive
    const flags = match.flags.includes('i') ? match.flags : match.flags + 'i';
    regex = new RegExp(match.source, flags);
    } else {
    regex = new RegExp(escapeForRegex(match), 'i');  // ← ADD 'i'
    }

    const result = regex.exec(fullText);
    if (!result) return;

    const start = result.index;
    const end = start + result[0].length - 1;

    if (style.bold) textElement.setBold(start, end, true);
    if (style.underline) textElement.setUnderline(start, end, true);
    if (style.italic) textElement.setItalic(start, end, true);
    if (style.color) textElement.setForegroundColor(start, end, style.color);
    }

/* ------------------ SHOW SIDEBAR ------------------ */
  // Main Function
    function showSidebar() {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle("Dr. I.E. Osadiaye's Medical Templates");
    DocumentApp.getUi().showSidebar(html);
    }

/* ------------------ EXPAND KEYWORDS FROM SIDEBAR ------------------ */
  // Main Function
    function expandKeywordsFromSidebar(keyword) {
    const body = DocumentApp.getActiveDocument().getBody();
    const normalized = keyword.toLowerCase().trim();

    // --- SPECIAL: Medication Table Trigger ---
    if (normalized === "/generatemedicinetable") {
    if (TABLE_ROW_BUFFER.length > 0) generateMedicineTableFromBuffer();
    insertDiagnosesIntoDocument();
    insertTemplatesIntoDocument();
    return;
    }

    // --- RESET HANDLING (match main engine behavior) ---
    if (normalized.endsWith("reset")) {
    const templateFn = TEMPLATE_DEFINITIONS[normalized];
    if (templateFn) {
    const template = templateFn();

    body.clear();
    insertTemplateAtIndex(body, template, 0);

    if (template.customAction) template.customAction();
    }
    return;
    }

    // --- NORMAL TEMPLATE HANDLING (BUFFER, DON'T INSERT) ---
    const templateFn = TEMPLATE_DEFINITIONS[normalized];
    if (templateFn) {
    const template = templateFn();


    // Default to 999 to ensure it goes to the very bottom unless specified otherwise
    let rank = template.rank || 999; 

    if (template.diagnoses && template.diagnoses.length > 0) {
    bufferDiagnoses(template.diagnoses);
    // Only override the rank if the diagnosis rank is HIGHER priority (smaller number)
    const firstKey = template.diagnoses[0];
    const diagRank = (DIAGNOSIS_REGISTRY[firstKey] && DIAGNOSIS_REGISTRY[firstKey].rank) || 999;
    rank = Math.min(rank, diagRank);
    }

    bufferTemplate(template, rank);

    if (template.customAction) template.customAction();
    }

    // --- MEDICATION COMMANDS ---
    if (normalized.startsWith("/c") && !TEMPLATE_DEFINITIONS[normalized]) {
    const medRow = processMedicationCommand(keyword);
    if (medRow) TABLE_ROW_BUFFER.push(medRow);
    }

    insertDiagnosesIntoDocument();
    insertTemplatesIntoDocument();

    // --- FINAL FLUSH (same as expandKeywords) ---
    if (TABLE_ROW_BUFFER.length > 0) generateMedicineTableFromBuffer();
    }
