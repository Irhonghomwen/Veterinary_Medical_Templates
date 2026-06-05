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


    function parseKeywordMetadata(rawKeyword) {
    let sex = "male"; // Default
    let plurality = "singular"; // Default
    let base = rawKeyword.toLowerCase();

    // Check for suffixes (order matters: check longer "females" before "female")
    if (base.endsWith("females")) {
      sex = "female"; plurality = "plural";
      base = base.replace(/females$/, "");
    } else if (base.endsWith("males")) {
      sex = "male"; plurality = "plural";
      base = base.replace(/males$/, "");
    } else if (base.endsWith("female")) {
      sex = "female"; plurality = "singular";
      base = base.replace(/female$/, "");
    } else if (base.endsWith("male")) {
      sex = "male"; plurality = "singular";
      base = base.replace(/male$/, "");
    }

    return { base, sex, plurality };
    }

  // Grammar Dictionary
    const GRAMMAR_DICTIONARY = {
    eyes: {
    singular: { a_cataract: "a complete cataract", a_corneal_ulcer: "A corneal ulcer", cataract: "cataract", cherry_eye: "a cherry eye", eye: "eye", gland: "gland", is: "is", it: "it", its: "its", make: "makes", this: "this", This: "This", ulcer: "ulcer", was: "was", pro: "it" },
    plural:   { a_cataract: "complete cataracts", a_corneal_ulcer: "Corneal ulcers", cataract: "cataracts", cherry_eye: "cherry eyes", eye: "eyes", gland: "glands", is: "are", it: "them", its: "their", make: "make", this: "these", This: "These", ulcer: "ulcers", was: "were", pro: "they" }
    },

    patella: {
    singular: { sub: "a luxating patella", bone: "kneecap", v: "slides", is: "is", pos: "its" },
    plural:   { sub: "luxating patellae", bone: "kneecaps", v: "slide", is: "are", pos: "their" }
    },

    wellness: {
    singular: { 
        begins: "begins", comes: "comes", Dogs: "Dog", dog: "dog", dogs: "dog's", dogss: "dog's", eats: "eats", gets: "gets",
        has: "has", he: "he", him: "him", his: "his", is: "is", life: "life", mother: "mother's", needs: "needs", puppy: "puppy",
        resists: "resists", round: "round", site: "site", shot: "shot", shows: "shows", steals: "steals", them: "them", was: "was", weighs: "weighs", 
         
    },
    plural: { 
        begins: "begin", comes: "come", Dogs: "Dogs", dog: "dogs", dogs: "dogs'", dogss: "dogs'", eats: "eat", gets: "get",
        has: "have", he: "they", him: "them", his: "their", is: "are", life: "lives", mother: "mothers'", needs: "need", puppy: "puppies",
        resists: "resist", round: "rounds", site: "sites", shot: "shots", shows: "show", steals: "steal", them: "them", was: "were", weighs: "weigh", 
          
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
  
/* ------------------ MEDICINE CABINET ------------------ */
  // Medications

    // Compiles a clean list of labels and keys for the sidebar to build buttons.
    function getMedicineRegistryKeys() {
    return Object.keys(MEDICINE_REGISTRY).map(key => {
    return {
    key: key,
    label: MEDICINE_REGISTRY[key].label
    };
    });
    }

    let TABLE_ROW_BUFFER = [];
    var MEDICINE_REGISTRY = {

    ADEQUANINJECTION: {
    label: "Adequan",
    instructions: "Injected in your dog’s muscle to improve joint health.",
    class: "Polysulfated glycosaminoglycan",
    sideEffects: "Well tolerated"
    },

    AMLODIPINE25: {
    label: "Amlodipine 2.5mg tablet",
    instructions: "Give your dog 1 tablet by mouth every 24 hours to manage high blood pressure. Recheck blood pressure in 1 week.",
    class: "Calcium channel blocker",
    sideEffects: "May cause lethargy, decreased appetite, or weight loss"
    },

    APOQUEL36: {
    label: "Apoquel 3.6mg (oclacitinib)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of allergies.",
    class: "Anti-allergy (JAK inhibitor)",
    sideEffects: "Over-suppresses the immune system when given with Zenrelia."
    },

    APOQUEL54: {
    label: "Apoquel 5.4mg (oclacitinib)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of allergies.",
    class: "Anti-allergy (JAK inhibitor)",
    sideEffects: "Over-suppresses the immune system when given with Zenrelia."
    },

    APOQUEL16: {
    label: "Apoquel 16mg (oclacitinib)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of allergies.",
    class: "Anti-allergy (JAK inhibitor)",
    sideEffects: "Over-suppresses the immune system when given with Zenrelia."
    },

    ATROPINE: {
    label: "Atropine 1% Ophthalmic Drops",
    instructions: "Starting today \nApply 1 drop in your dog’s affected eye to ease pain. Apply BEFORE eye ointments.",
    class: "Anticholinergic",
    sideEffects: "May cause light sensitivity"
    },

    BENAZEPRIL5: {
    label: "Benazepril 5mg",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for management of heart failure.",
    class: "Angiotensin converting enzyme (ACE) inhibitor",
    sideEffects: "May cause bloodwork abnormalities and hypotension"
    },

    CARPROFEN25: {
    label: "Carprofen 25mg",
    instructions: "Give your dog 1 tablet by mouth every 12 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite. DO NOT USE WITHIN 3 DAYS OF OTHER NSAIDs OR STEROIDS."
    },

    CARPROFEN75: {
    label: "Carprofen 75mg",
    instructions: "Give your dog 1 tablet by mouth every 12 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite. DO NOT USE WITHIN 3 DAYS OF OTHER NSAIDs OR STEROIDS."
    },

    CARPROFEN100: {
    label: "Carprofen 100mg",
    instructions: "Give your dog 1 tablet by mouth every 12 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite. DO NOT USE WITHIN 3 DAYS OF OTHER NSAIDs OR STEROIDS."
    },

    CEFPODOXIME100: {
    label: "Cefpodoxime 100mg",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite."
    },

    CEFPODOXIME200: {
    label: "Cefpodoxime 200mg",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite."
    },

    CLAVAMOX625: {
    label: "Amoxicillin clavulanate 62.5mg",
    instructions: "Give your dog 1 tablet by mouth every 12 hours for treatment of infection.",
    class: "Broad spectrum potentiated antibiotic",
    sideEffects: "Vomiting, diarrhea, or decreased appetite (less common if given with a meal)."
    },

    CLAVAMOX125: {
    label: "Amoxicillin clavulanate 125mg",
    instructions: "Give your dog 1 tablet by mouth every 12 hours for treatment of infection.",
    class: "Broad spectrum potentiated antibiotic",
    sideEffects: "Vomiting, diarrhea, or decreased appetite (less common if given with a meal)."
    },

    CLAVAMOX250: {
    label: "Amoxicillin clavulanate 250mg",
    instructions: "Give your dog 1 tablet by mouth every 12 hours for treatment of infection.",
    class: "Broad spectrum potentiated antibiotic",
    sideEffects: "Vomiting, diarrhea, or decreased appetite (less common if given with a meal)."
    },

    CLAVAMOX375: {
    label: "Amoxicillin clavulanate 375mg",
    instructions: "Give your dog 1 tablet by mouth every 12 hours for treatment of infection.",
    class: "Broad spectrum potentiated antibiotic",
    sideEffects: "Vomiting, diarrhea, or decreased appetite (less common if given with a meal)."
    },

    COUGHTABLETS: {
    label: "Cough tablets\n(guaifenesin & dextromethorphan hydrobromide)",
    instructions: "Give your dog 1 tablet by mouth every 4 - 6 hours for suppression of cough",
    class: "Antitussive (cough suppressant)",
    sideEffects: "Well tolerated"
    },

    CYTOPOINTINJECTION: {
    label: "Cytopoint injection (lokivetmab)",
    instructions: "Medication injected beneath your dog’s skin to control allergies over the next 4 - 8 weeks.",
    class: "Anti-allergy (monoclonal antibody)",
    sideEffects: "Well tolerated"
    },

    DORZOLAMIDE: {
    label: "Dorzolamide",
    instructions: "Apply 1 drop in each eye every 8 hours for management of glaucoma.",
    class: "Carbonic anhydrase inhibitor",
    sideEffects: "May cause ocular discomfort. Medication tastes bitter and some pets may make faces due to this."
    },

    DOXYCYCLINE100: {
    label: "Doxycycline 100mg",
    instructions: "Give your dog 1 tablet by mouth with food every 12 hours for treatment of bacterial infection. Give until gone.",
    class: "Antibiotic",
    sideEffects: "May cause vomiting or diarrhea."
    },

    DOXYCYCLINE200: {
    label: "Doxycycline 100mg",
    instructions: "Give your dog 1 tablet by mouth with food every 12 hours for treatment of bacterial infection. Give until gone.",
    class: "Antibiotic",
    sideEffects: "May cause vomiting or diarrhea."
    },

    FUROSEMIDE125: {
    label: "Furosemide 12.5 mg",
    instructions: "Give your dog 1 tablet by mouth every 8 - 12 hours to drain fluid from the lungs.",
    class: "Diuretic",
    sideEffects: "May cause increased drinking and urination or bloodwork abnormalities"
    },

    FUROSEMIDE20: {
    label: "Furosemide 20 mg",
    instructions: "Give your dog 1 tablet by mouth every 8 - 12 hours to drain fluid from the lungs.",
    class: "Diuretic",
    sideEffects: "May cause increased drinking and urination or bloodwork abnormalities"
    },

    FUROSEMIDE50: {
    label: "Furosemide 50 mg",
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

    GABAPENTIN50: {
    label: "Gabapentin 50mg",
    instructions: "Give your dog 1 tablet by mouth every 8 - 12 hours for treatment of pain.",
    class: "Analgesia, sedative",
    sideEffects: "May cause sedation"
    },

    GABAPENTIN100: {
    label: "Gabapentin 100mg",
    instructions: "Give your dog 1 capsule by mouth every 8 - 12 hours for treatment of pain.",
    class: "Analgesia, sedative",
    sideEffects: "May cause sedation"
    },

    GABAPENTIN200: {
    label: "Gabapentin 200mg",
    instructions: "Give your dog 1 tablet by mouth every 8 - 12 hours for treatment of pain.",
    class: "Analgesia, sedative",
    sideEffects: "May cause sedation"
    },

    GABAPENTIN300: {
    label: "Gabapentin 300mg",
    instructions: "Give your dog 1 capsule by mouth every 8 - 12 hours for treatment of pain.",
    class: "Analgesia, sedative",
    sideEffects: "May cause sedation"
    },

    GABAPENTIN400: {
    label: "Gabapentin 400mg",
    instructions: "Give your dog 1 tablet by mouth every 8 - 12 hours for treatment of pain.",
    class: "Analgesia, sedative",
    sideEffects: "May cause sedation"
    },

    GABAPENTIN600: {
    label: "Gabapentin 600mg",
    instructions: "Give your dog 1 tablet by mouth every 8 - 12 hours for treatment of pain.",
    class: "Analgesia, sedative",
    sideEffects: "May cause sedation"
    },

    GRAPIPRANT20: {
    label: "Galliprant 20mg (grapiprant)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite. DO NOT USE WITHIN 3 DAYS OF OTHER NSAIDs OR STEROIDS."
    },

    GRAPIPRANT60: {
    label: "Galliprant 60mg (grapiprant)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite. DO NOT USE WITHIN 3 DAYS OF OTHER NSAIDs OR STEROIDS."
    },

    GRAPIPRANT100: {
    label: "Galliprant 100mg (grapiprant)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite. DO NOT USE WITHIN 3 DAYS OF OTHER NSAIDs OR STEROIDS."
    },

    HEARTGARD25: {
    label: "Heartgard Plus Blue 0 - 25 lbs",
    instructions: "Give your dog 1 chewable tablet every 30 days for prevention of heartworms and common intestinal parasites.",
    class: "Parasiticide",
    sideEffects: "Rarely causes vomiting or diarrhea"
    },

    HEARTGARD50: {
    label: "Heartgard Plus Green 25 - 50 lbs",
    instructions: "Give your dog 1 chewable tablet every 30 days for prevention of heartworms and common intestinal parasites.",
    class: "Parasiticide",
    sideEffects: "Rarely causes vomiting or diarrhea"
    },

    HEARTGARD100: {
    label: "Heartgard Plus Brown 50 - 100 lbs",
    instructions: "Give your dog 1 chewable tablet every 30 days for prevention of heartworms and common intestinal parasites.",
    class: "Parasiticide",
    sideEffects: "Rarely causes vomiting or diarrhea"
    },

    HILLSGIBIOMECANNED: {
    label: "Hill's Prescription Gastrointestinal Biome (canned food)",
    instructions: "Feed your dog 1 can by mouth every 12 hours until all cans are gone for treatment of diarrhea.",
    class: "Prebiotic, probiotic, & postbiotic",
    sideEffects: "Well tolerated"
    },

    HILLSGIBIOMEDRY: {
    label: "Hill's Prescription Gastrointestinal Biome (dry food)",
    instructions: "Feed your dog 1 cup by mouth every 12 hours until all cans are gone for treatment of diarrhea.",
    class: "Prebiotic, probiotic, & postbiotic",
    sideEffects: "Well tolerated"
    },

    HILLSIDCANNED: {
    label: "Hill's Prescription i/d (canned food)",
    instructions: "Feed your dog 1 can by mouth every 12 hours until all cans are gone for treatment of diarrhea.",
    class: "Prebiotic, probiotic, & postbiotic",
    sideEffects: "Well tolerated"
    },

    HILLSIDDRY: {
    label: "Hill's Prescription i/d Biome (dry food)",
    instructions: "Feed your dog 1 cup by mouth every 12 hours until all cans are gone for treatment of diarrhea.",
    class: "Prebiotic, probiotic, & postbiotic",
    sideEffects: "Well tolerated"
    },

    LEVOTHYROXINE01: {
    label: "Thyro-tabs 0.1 mg (levothyroxine)",
    instructions: "Give your dog 1 tablet by mouth without food every 12 hours. Recheck thyroid levels in 4 - 6 weeks.",
    class: "Thyroid hormone",
    sideEffects: "Well tolerated"
    },

    LEVOTHYROXINE02: {
    label: "Thyro-tabs 0.2 mg (levothyroxine)",
    instructions: "Give your dog 1 tablet by mouth without food every 12 hours. Recheck thyroid levels in 4 - 6 weeks.",
    class: "Thyroid hormone",
    sideEffects: "Well tolerated"
    },

    LEVOTHYROXINE03: {
    label: "Thyro-tabs 0.3 mg (levothyroxine)",
    instructions: "Give your dog 1 tablet by mouth without food every 12 hours. Recheck thyroid levels in 4 - 6 weeks.",
    class: "Thyroid hormone",
    sideEffects: "Well tolerated"
    },

    LEVOTHYROXINE04: {
    label: "Thyro-tabs 0.4 mg (levothyroxine)",
    instructions: "Give your dog 1 tablet by mouth without food every 12 hours. Recheck thyroid levels in 4 - 6 weeks.",
    class: "Thyroid hormone",
    sideEffects: "Well tolerated"
    },

    LEVOTHYROXINE05: {
    label: "Thyro-tabs 0.5 mg (levothyroxine)",
    instructions: "Give your dog 1 tablet by mouth without food every 12 hours. Recheck thyroid levels in 4 - 6 weeks.",
    class: "Thyroid hormone",
    sideEffects: "Well tolerated"
    },

    LEVOTHYROXINE06: {
    label: "Thyro-tabs 0.6 mg (levothyroxine)",
    instructions: "Give your dog 1 tablet by mouth without food every 12 hours. Recheck thyroid levels in 4 - 6 weeks.",
    class: "Thyroid hormone",
    sideEffects: "Well tolerated"
    },

    LEVOTHYROXINE07: {
    label: "Thyro-tabs 0.7 mg (levothyroxine)",
    instructions: "Give your dog 1 tablet by mouth without food every 12 hours. Recheck thyroid levels in 4 - 6 weeks.",
    class: "Thyroid hormone",
    sideEffects: "Well tolerated"
    },

    LEVOTHYROXINE08: {
    label: "Thyro-tabs 0.8 mg (levothyroxine)",
    instructions: "Give your dog 1 tablet by mouth without food every 12 hours. Recheck thyroid levels in 4 - 6 weeks.",
    class: "Thyroid hormone",
    sideEffects: "Well tolerated"
    },

    LEVOTHYROXINE1: {
    label: "Thyro-tabs 1.0 mg (levothyroxine)",
    instructions: "Give your dog 1 tablet by mouth without food every 12 hours. Recheck thyroid levels in 4 - 6 weeks.",
    class: "Thyroid hormone",
    sideEffects: "Well tolerated"
    },

    LIBRELAINJECTION: {
    label: "Librela injection (bedinvetmab)",
    instructions: "Medication injected beneath your dog’s skin to control arthritis over the next 4 weeks.",
    class: "Anti-arthritis (monoclonal antibody)",
    sideEffects: "Rarely causes seizures, urinary tract infection, or skin infections."
    },

    MAROPITANT16: {
    label: "Cerenia 16mg (maropitant)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of vomiting & nausea.",
    class: "Antiemetic",
    sideEffects: "Well tolerated"
    },

    MAROPITANT24: {
    label: "Cerenia 24mg (maropitant)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of vomiting & nausea.",
    class: "Antiemetic",
    sideEffects: "Well tolerated"
    },

    MAROPITANT60: {
    label: "Cerenia 60mg (maropitant)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of vomiting & nausea.",
    class: "Antiemetic",
    sideEffects: "Well tolerated"
    },

    MAROPITANT160: {
    label: "Cerenia 160mg (maropitant)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of vomiting & nausea.",
    class: "Antiemetic",
    sideEffects: "Well tolerated"
    },

    MAROPITANTINJECTION: {
    label: "Cerenia injection (maropitant)",
    instructions: "Medication injected beneath your dog’s skin to control nausea & vomiting.",
    class: "Antiemetic",
    sideEffects: "Well tolerated"
    },

    MELOXICAM75: {
    label: "Meloxicam 7.5mg",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite. DO NOT USE WITH OTHER NSAIDs OR STEROIDS."
    },

    MELOXICAM15LIQUID: {
    label: "Meloxicam liquid 1.5mg/mL",
    instructions: "Give your dog 1 mL by mouth every 24 hours for pain and inflammation.",
    class: "Non-steroidal anti-inflammatory drug (NSAID)",
    sideEffects: "Vomiting, diarrhea, or decreased appetite. DO NOT USE WITHIN 3 DAYS OF OTHER NSAIDs OR STEROIDS."
    },

    NEOPOLYBACOINTMENT: {
    label: "NeoPolyBac ointment (neomycin, polymyxin, bacitracin)",
    instructions: "Apply ¼ inch strip in your dog’s affected eye every 8 - 12 hours for treatment of infection & inflammation.",
    class: "Antibiotic, anti-inflammatory",
    sideEffects: "Well tolerated"
    },

    NEOPOLYBACHYDROOINTMENT: {
    label: "NeoPolyBac with Hydrocortisone ointment\n(neomycin, polymyxin, bacitracin, hydrocortisone)",
    instructions: "Apply ¼ inch strip in the affected eye every 8 - 12 hours for treatment of corneal ulcer.",
    class: "Antibiotic, anti-inflammatory",
    sideEffects: "Well tolerated"
    },

    NEOPOLYDEXSUSPENSION: {
    label: "NeoPolyDex Suspension (neomycin, polymyxin B, dexamethasone)",
    instructions: "Apply 1 - 2 drops in your dog’s affected eye every 8 - 12 hours for treatment of infection & inflammation.",
    class: "Antibiotic, anti-inflammatory",
    sideEffects: "Well tolerated"
    },
    
    NEXGARD1: {
    label: "Nexgard Orange 4 - 10 lbs",
    instructions: "Give your dog 1 chewable tablet every 30 days for prevention of fleas and ticks.",
    class: "Parasiticide",
    sideEffects: "Rarely causes vomiting or diarrhea"
    },

    NEXGARD2: {
    label: "Nexgard Blue 10.1 - 24 lbs",
    instructions: "Give your dog 1 chewable tablet every 30 days for prevention of fleas and ticks.",
    class: "Parasiticide",
    sideEffects: "Rarely causes vomiting or diarrhea"
    },

    NEXGARD3: {
    label: "Nexgard Purple 24.1 - 60 lbs",
    instructions: "Give your dog 1 chewable tablet every 30 days for prevention of fleas and ticks.",
    class: "Parasiticide",
    sideEffects: "Rarely causes vomiting or diarrhea"
    },

    NEXGARD4: {
    label: "Nexgard Red 60.1 - 121",
    instructions: "Give your dog 1 chewable tablet every 30 days for prevention of fleas and ticks.",
    class: "Parasiticide",
    sideEffects: "Rarely causes vomiting or diarrhea"
    },

    ONDASETRON4: {
    label: "Ondansetron 4mg",
    instructions: "Give your dog 1 tablet every 8 - 12 hours for management of pancreatitis.",
    class: "5-HT3 Antagonist, antiemetic",
    sideEffects: "Well tolerated"
    },

    ONDASETRON8: {
    label: "Ondansetron 8mg",
    instructions: "Give your dog 1 tablet every 8 - 12 hours for management of pancreatitis.",
    class: "5-HT3 Antagonist, antiemetic",
    sideEffects: "Well tolerated"
    },

    OPTIMMUNEOINTMENT: {
    label: "Optimmune ointment (cyclosporine 0.2%)",
    instructions: "Apply ¼ inch in your dog’s affected eye every 8 hours for treatment of dry eye. Apply 5 minutes AFTER other eye drop medicine.",
    class: "Immunosuppressant",
    sideEffects: "Well tolerated"
    },

    OPTIXCARE: {
    label: "Optixcare Eye Lube",
    instructions: "Apply ¼ inch or 1 - 2 drops in your dog’s affected eye every 8 hours for treatment of dry eye. Apply 5 minutes AFTER other eye drop medicine.",
    class: "Lubricant",
    sideEffects: "Well tolerated"
    },

    PIMOBENDAN125: {
    label: "Vetmedin 1.25mg\n(pimobendan)",
    instructions: "Give your dog 1 tablet by mouth every 12 hours to increase heart contractility & function.",
    class: "Inotropic agent",
    sideEffects: "Rarely causes vomiting (less than 1% of dogs)"
    },
    
    PIMOBENDAN25: {
    label: "Vetmedin 2.5mg\n(pimobendan)",
    instructions: "Give your dog 1 tablet by mouth every 12 hours to increase heart contractility & function.",
    class: "Inotropic agent",
    sideEffects: "Rarely causes vomiting (less than 1% of dogs)"
    },

    PIMOBENDAN5: {
    label: "Vetmedin 5mg\n(pimobendan)",
    instructions: "Give your dog 1 tablet by mouth every 12 hours to increase heart contractility & function.",
    class: "Inotropic agent",
    sideEffects: "Rarely causes vomiting (less than 1% of dogs)"
    },

    PIMOBENDAN10: {
    label: "Vetmedin 10mg\n(pimobendan)",
    instructions: "Give your dog 1 tablet by mouth every 12 hours to increase heart contractility & function.",
    class: "Inotropic agent",
    sideEffects: "Rarely causes vomiting (less than 1% of dogs)"
    },

    PREDNISOLONE5: {
    label: "Prednisolone 5mg",
    instructions: "Give 1 tablet by mouth every 12 hours for 7 days, then 1 tablet every 24 hours for 7 days, then 1 every 48 hours for a total of 8 times, then discontinue.",
    class: "Corticosteroid",
    sideEffects: "May cause vomiting, diarrhea, increased appetite, and/or increased drinking & urination"
    },

    PREDNISOLONE10: {
    label: "Prednisolone 10mg",
    instructions: "Give 1 tablet by mouth every 12 hours for 7 days, then 1 tablet every 24 hours for 7 days, then 1 every 48 hours for a total of 8 times, then discontinue.",
    class: "Corticosteroid",
    sideEffects: "May cause vomiting, diarrhea, increased appetite, and/or increased drinking & urination"
    },

    PREDNISOLONE20: {
    label: "Prednisolone 20mg",
    instructions: "Give 1 tablet by mouth every 12 hours for 7 days, then 1 tablet every 24 hours for 7 days, then 1 every 48 hours for a total of 8 times, then discontinue.",
    class: "Corticosteroid",
    sideEffects: "May cause vomiting, diarrhea, increased appetite, and/or increased drinking & urination"
    },

    PREDNISONE5: {
    label: "Prednisone 5mg",
    instructions: "Give 1 tablet by mouth every 12 hours for 7 days, then 1 tablet every 24 hours for 7 days, then 1 every 48 hours for a total of 8 times, then discontinue.",
    class: "Corticosteroid",
    sideEffects: "May cause vomiting, diarrhea, increased appetite, and/or increased drinking & urination"
    },

    PREDNISONE10: {
    label: "Prednisone 10mg",
    instructions: "Give 1 tablet by mouth every 12 hours for 7 days, then 1 tablet every 24 hours for 7 days, then 1 every 48 hours for a total of 8 times, then discontinue.",
    class: "Corticosteroid",
    sideEffects: "May cause vomiting, diarrhea, increased appetite, and/or increased drinking & urination"
    },

    PREDNISONE20: {
    label: "Prednisone 20mg",
    instructions: "Give 1 tablet by mouth every 12 hours for 7 days, then 1 tablet every 24 hours for 7 days, then 1 every 48 hours for a total of 8 times, then discontinue.",
    class: "Corticosteroid",
    sideEffects: "May cause vomiting, diarrhea, increased appetite, and/or increased drinking & urination"
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

    PROVIABLECAPSULES: {
    label: "Proviable Forte capsules",
    instructions: "Give your dog 1 capsule by mouth every 24 hours for 15 days to treat diarrhea.",
    class: "Probiotic",
    sideEffects: "Well tolerated"
    },

    PROVIABLEPASTE: {
    label: "Proviable Forte paste",
    instructions: "Give your dog 1 mL by mouth every 8 hours. Give for 3 days or until diarrhea stops, whichever occurs first.",
    class: "Antidiarrheal",
    sideEffects: "Well tolerated"
    },

    PURINAENCANNED: {
    label: "Purina Pro Plan Gastroenteric Diet (EN)",
    instructions: "Feed your dog 1 can by mouth every 12 hours until all cans are gone for treatment of diarrhea.",
    class: "Prebiotic, probiotic, & postbiotic",
    sideEffects: "Well tolerated"
    },

    PURINAENLOWFATCANNED: {
    label: "Purina Pro Plan Gastroenteric Diet (EN) - low fat wet food",
    instructions: "Feed your dog 1 can by mouth every 12 hours until all cans are gone for treatment of diarrhea.",
    class: "Prebiotic, probiotic, & postbiotic",
    sideEffects: "Well tolerated"
    },

    PURINAENDRY: {
    label: "Purina Pro Plan Gastroenteric Diet (EN) dry food",
    instructions: "Feed your dog 1 cup by mouth every 12 hours until all cans are gone for treatment of diarrhea.",
    class: "Prebiotic, probiotic, & postbiotic",
    sideEffects: "Well tolerated"
    },

    PURINAENLOWFATDRY: {
    label: "Purina Pro Plan Gastroenteric Diet (EN) - low fat dry food",
    instructions: "Feed your dog 1 cup by mouth every 12 hours until all cans are gone for treatment of diarrhea.",
    class: "Prebiotic, probiotic, & postbiotic",
    sideEffects: "Well tolerated"
    },

    REVOLUTION: {
    label: "Revolution",
    instructions: "Apply contents between your dog’s ears every 30 days for prevention of heartworms, fleas, ticks, and common intestinal parasites.",
    class: "Parasiticide",
    sideEffects: "Rarely causes redness of the skin or hair loss"
    },

    REVOLUTION1: {
    label: "Revolution 0 - 5 lbs",
    instructions: "Apply contents between your dog’s ears every 30 days for prevention of heartworms, fleas, ticks, and common intestinal parasites.",
    class: "Parasiticide",
    sideEffects: "Rarely causes redness of the skin or hair loss"
    },

    REVOLUTION2: {
    label: "Revolution 5.1 - 10 lbs",
    instructions: "Apply contents between your dog’s ears every 30 days for prevention of heartworms, fleas, ticks, and common intestinal parasites.",
    class: "Parasiticide",
    sideEffects: "Rarely causes redness of the skin or hair loss"
    },

    REVOLUTION3: {
    label: "Revolution 10.1 - 20 lbs ",
    instructions: "Apply contents between your dog’s ears every 30 days for prevention of heartworms, fleas, ticks, and common intestinal parasites.",
    class: "Parasiticide",
    sideEffects: "Rarely causes redness of the skin or hair loss"
    },

    REVOLUTION4: {
    label: "Revolution 20.1 - 40 lbs",
    instructions: "Apply contents between your dog’s ears every 30 days for prevention of heartworms, fleas, ticks, and common intestinal parasites.",
    class: "Parasiticide",
    sideEffects: "Rarely causes redness of the skin or hair loss"
    },

    REVOLUTION5: {
    label: "Revolution 40.1 - 85 lbs",
    instructions: "Apply contents between your dog’s ears every 30 days for prevention of heartworms, fleas, ticks, and common intestinal parasites.",
    class: "Parasiticide",
    sideEffects: "Rarely causes redness of the skin or hair loss"
    },

    REVOLUTION6: {
    label: "Revolution 85.1 - 130 lbs",
    instructions: "Apply contents between your dog’s ears every 30 days for prevention of heartworms, fleas, ticks, and common intestinal parasites.",
    class: "Parasiticide",
    sideEffects: "Rarely causes redness of the skin or hair loss"
    },

    SIMPARICATRIO1: {
    label: "Simparica Trio 2.8 - 5.5 lbs",
    instructions: "Give your dog 1 chewable tablet by mouth every 30 days for prevention of heartworms, fleas, ticks, & common intestinal parasites.",
    class: "Antiparasitic",
    sideEffects: "Rarely causes vomiting, diarrhea, or neurologic abnormalities"
    },

    SIMPARICATRIO2: {
    label: "Simparica Trio 5.6 - 11.0 lbs",
    instructions: "Give your dog 1 chewable tablet by mouth every 30 days for prevention of heartworms, fleas, ticks, & common intestinal parasites.",
    class: "Antiparasitic",
    sideEffects: "Rarely causes vomiting, diarrhea, or neurologic abnormalities"
    },

    SIMPARICATRIO3: {
    label: "Simparica Trio 11.1 - 22.0 lbs",
    instructions: "Give your dog 1 chewable tablet by mouth every 30 days for prevention of heartworms, fleas, ticks, & common intestinal parasites.",
    class: "Antiparasitic",
    sideEffects: "Rarely causes vomiting, diarrhea, or neurologic abnormalities"
    },

    SIMPARICATRIO4: {
    label: "Simparica Trio 22.1 - 44.0 lbs",
    instructions: "Give your dog 1 chewable tablet by mouth every 30 days for prevention of heartworms, fleas, ticks, & common intestinal parasites.",
    class: "Antiparasitic",
    sideEffects: "Rarely causes vomiting, diarrhea, or neurologic abnormalities"
    },

    SIMPARICATRIO5: {
    label: "Simparica Trio 44.1 - 88.0 lbs",
    instructions: "Give your dog 1 chewable tablet by mouth every 30 days for prevention of heartworms, fleas, ticks, & common intestinal parasites.",
    class: "Antiparasitic",
    sideEffects: "Rarely causes vomiting, diarrhea, or neurologic abnormalities"
    },

    SIMPARICATRIO6: {
    label: "Simparica Trio 88.1 - 132.0 lbs",
    instructions: "Give your dog 1 chewable tablet by mouth every 30 days for prevention of heartworms, fleas, ticks, & common intestinal parasites.",
    class: "Antiparasitic",
    sideEffects: "Rarely causes vomiting, diarrhea, or neurologic abnormalities"
    },

    SPIRONOLACTONE25: {
    label: "Spironolactone 25mg",
    instructions: "Give your dog 1 tablet by mouth every 24 hours to decrease heart workload.",
    class: "Diuretic",
    sideEffects: "May cause bloodwork abnormalities (elevated BUN)"
    },

    SUCRALFATE: {
    label: "Sucralfate 1 gram",
    instructions: "Crush/dissolve 1 tablet with 3 - 5 mL of water before giving by mouth every 12 hours. GIVE 2 HOURS BEFORE OR AFTER OTHER FOODS OR MEDS.",
    class: "Antiulcer",
    sideEffects: "May cause decreased absorption of other medicines & food given by mouth"
    },

    SUBCUTANEOUSFLUIDS: {
    label: "Subcutaneous fluids",
    instructions: "Fluids injected beneath your dog’s skin to rehydrate patient",
    class: "Fluids",
    sideEffects: "Well tolerated"
    },

    SYNOTIC: {
    label: "Synotic Otic Solution",
    instructions: "Starting today\nApply up to 5 drops in your dog’s affected ear every 12 hours for 1 week, then discontinue.",
    class: "Corticosteroid",
    sideEffects: "May cause short term ear discomfort or increased thirst/urination." },

    TACROLIMUS: {
    label: "Tacrolimus 0.02% ophthalmic solution",
    instructions: "Apply 1 - 2 drops in your dog’s affected eye every 12 hours for management of dry eye. Apply 5 minutes BEFORE other eye drop medicine.",
    class: "Immunosuppressant",
    sideEffects: "Well tolerated"
    },

    TRAZODONE50: {
    label: "Trazodone 50mg",
    instructions: "Give your dog 1 tablet the night before and 2 hours prior to a morning appointment.",
    class: "Anxiolytic",
    sideEffects: "May cause sedation or hyperactivity"
    },

    TRAZODONE100: {
    label: "Trazodone 100mg",
    instructions: "Give your dog 1 tablet the night before and 2 hours prior to a morning appointment.",
    class: "Anxiolytic",
    sideEffects: "May cause sedation or hyperactivity"
    },

    TRAZODONE150: {
    label: "Trazodone 150mg",
    instructions: "Give your dog 1 tablet the night before and 2 hours prior to a morning appointment.",
    class: "Anxiolytic",
    sideEffects: "May cause sedation or hyperactivity"
    },

    TRILOSTANE5: {
    label: "Vetoryl 5mg (trilostane)",
    instructions: "Give your dog 1 capsule by mouth every 24 hours for management of hyperadrenocorticism.",
    class: "Adrenal suppressant",
    sideEffects: "May cause vomiting, decreased appetite, or lethargy."
    },

    TRILOSTANE10: {
    label: "Vetoryl 10mg (trilostane)",
    instructions: "Give your dog 1 capsule by mouth every 24 hours for management of hyperadrenocorticism.",
    class: "Adrenal suppressant",
    sideEffects: "May cause vomiting, decreased appetite, or lethargy."
    },

    TRILOSTANE20: {
    label: "Vetoryl 20mg (trilostane)",
    instructions: "Give your dog 1 capsule by mouth every 24 hours for management of hyperadrenocorticism.",
    class: "Adrenal suppressant",
    sideEffects: "May cause vomiting, decreased appetite, or lethargy."
    },

    TRILOSTANE30: {
    label: "Vetoryl 30mg (trilostane)",
    instructions: "Give your dog 1 capsule by mouth every 24 hours for management of hyperadrenocorticism.",
    class: "Adrenal suppressant",
    sideEffects: "May cause vomiting, decreased appetite, or lethargy."
    },

    TRILOSTANE60: {
    label: "Vetoryl 60mg (trilostane)",
    instructions: "Give your dog 1 capsule by mouth every 24 hours for management of hyperadrenocorticism.",
    class: "Adrenal suppressant",
    sideEffects: "May cause vomiting, decreased appetite, or lethargy."
    },

    TRILOSTANE120: {
    label: "Vetoryl 120mg (trilostane)",
    instructions: "Give your dog 1 capsule by mouth every 24 hours for management of hyperadrenocorticism.",
    class: "Adrenal suppressant",
    sideEffects: "May cause vomiting, decreased appetite, or lethargy."
    },

    VETSULIN: {
    label: "Vetsulin 40 units/mL (porcine insulin zinc)",
    instructions: "Give your dog 1 unit every 12 hours. Give after eating. You can skip an injection once if your dog doesn’t eat breakfast/dinner.",
    class: "Hormone",
    sideEffects: "May cause low blood sugar (lethargy, drunken appearance, or seizures)"
    },


    ZENRELIA48: {
    label: "Zenrelia 4.8mg (ilunocitinib)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of allergies.",
    class: "Anti-allergy (JAK inhibitor)",
    sideEffects: "Well tolerated"
    },

    ZENRELIA64: {
    label: "Zenrelia 6.4mg (ilunocitinib)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of allergies.",
    class: "Anti-allergy (JAK inhibitor)",
    sideEffects: "Well tolerated"
    },

    ZENRELIA85: {
    label: "Zenrelia 8.5mg (ilunocitinib)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of allergies.",
    class: "Anti-allergy (JAK inhibitor)",
    sideEffects: "Well tolerated"
    },

    ZENRELIA15: {
    label: "Zenrelia 15mg (ilunocitinib)",
    instructions: "Give your dog 1 tablet by mouth every 24 hours for treatment of allergies.",
    class: "Anti-allergy (JAK inhibitor)",
    sideEffects: "Well tolerated"
    },
    };

    function insertMedicineFromSidebar(medKey, prefixKey) {
    // 1. Fallback safety check if the base medication key doesn't exist
    if (!MEDICINE_REGISTRY[medKey]) {
    throw new Error("Medication key '" + medKey + "' was not found in the Registry.");
    }

    // 2. Set default fallback configuration if no prefix was passed over
    if (!prefixKey) {
    prefixKey = "START";
    }

    // 3. Construct the precise target uppercase lookup token matched in your Forward Generator
    const targetLookupKey = (medKey + prefixKey).toUpperCase();
    const commandData = MED_COMMAND_LOOKUP[targetLookupKey];

    if (!commandData) {
    throw new Error("Could not find configuration data for combination key: " + targetLookupKey);
    }

    // 4. Send metadata payload into your active template routing engine
    TABLE_ROW_BUFFER.push(commandData);

    // Natively triggers sorting, deduplication, row background coloring, and first-line bold/underlines
    generateMedicineTableFromBuffer();
    }

  // Medication Prefixes & Registry Logic
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
    CONTINUE: "#B6D7A8",    // green
    DISCONTINUE: "#EA9999",  // red
    ASNEEDED: "#FFE599",    // yellow
    NEWDOSE: "#B4A7D6",      // purple
    CLINIC: "#A4C2F4"        // blue
    };

    // Helper for the Reverse Generator to clean up scraped text
    function stripMedPrefix(text) {
    if (!text) return "";
    let cleanText = text.trim();

    for (let key in MED_PREFIX) {
    let prefixText = MED_PREFIX[key];
    // Check if the instruction starts with the prefix text
    if (cleanText.startsWith(prefixText)) {
    // Remove prefix and any immediate newlines/spaces that follow it
    cleanText = cleanText.substring(prefixText.length).replace(/^[\n\r\s]+/, "");
    break; 
    }
    }
    return cleanText;
    }

    // Precompute every valid medication command (Forward Generation)
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

    // Default START command
    const defaultKey = (medKey + "START").toUpperCase();
    MED_COMMAND_LOOKUP[medKey.toUpperCase()] = MED_COMMAND_LOOKUP[defaultKey];
    });

    // Medication Command Processor
    function processMedicationCommand(keyword) {
    const cmd = keyword.replace(/^\/c/i, "").toUpperCase().trim();
    return MED_COMMAND_LOOKUP[cmd] || null;
    }

/* ------------------ DIAGNOSIS & TEMPLATE BUFFER/RANKING ------------------ */
  // Diagnosis Registry & Rank
    const DIAGNOSIS_REGISTRY = {
      // Examples
      // 0 - 99: Top priority, emergency, disregard all else (heart murmur, heart failure, etc.)
      // 100 - 199: Quality of life (arthritis)
      // 200 - 299: High priority, recheck & preventative care (corneal ulcer, glaucoma, hypertension)
      // 300 - 399: Medium-high priority, treatment advised (periodontal disease, otitis externa, etc.)
      // 400 - 499: Medium priority, lifelong management (atopic dermatitis, BOAS, bronchitis)
      // 500 - 599: Medium-low priority, treatment advised (conjunctivitis, bordetellosis, etc.)
      // 600 - 699: Low-high priority, medical attention (2nd degree AV block)
      // 700 - 799: Low-medium priority, client attention (overweight, blind, collapsing trachea)
      // 800 - 899: Low priority, advise (underweight)
      // 900 - 999: Non-vital, incidental findings, no treatment necessary (nuclear sclerosis)

    ACUTE_GASTROENTERITIS_DIARRHEA_ONLY: {
    text: "Acute gastroenteritis (diarrhea)",
    rank: 500
    },

    ACUTE_GASTROENTERITIS_VOMITING_ONLY: {
    text: "Acute gastroenteritis (vomiting)",
    rank: 500
    },

    ACUTE_GASTROENTERITIS_VOMITING_DIARRHEA: {
    text: "Acute gastroenteritis (vomiting & diarrhea)",
    rank: 500
    },

    ATOPIC_DERMATITIS: {
    text: "Atopic dermatitis (allergies)",
    rank: 400
    },

    BLIND: {
    text: "Blind",
    rank: 710
    },
    
    BORDETELLOSIS_PRESUMED: {
    text: "Bordetellosis (presumed)",
    rank: 510
    },

    BRACHYCEPHALIC_OBSTRUCTIVE_AIRWAY_SYNDROME: {
    text: "Brachycephalic obstructive airway syndrome (BOAS)",
    rank: 400
    },

    CATARACTS: {
    text: "Cataracts",
    rank: 710
    },

    CHERRY_EYE: {
    text: "Cherry eye",
    rank: 321
    },

    CHRONIC_BRONCHITIS_PRESUMED: {
    text: "Chronic bronchitis (presumed)",
    rank: 430
    },

    COLLAPSING_TRACHEA: {
    text: "Collapsing trachea",
    rank: 701
    },

    CONJUNCTIVITIS_DIAGNOSED: {
    text: "Conjunctivitis",
    rank: 580
    },

    CONJUNCTIVITIS_PRESUMED: {
    text: "Conjunctivitis (presumed)",
    rank: 780
    },

    CORNEAL_ULCER: {
    text: "Corneal ulcer",
    rank: 220
    },

    DIABETES_MELLITUS: {
    text: "Diabetes mellitus",
    rank: 4
    },

    ENTROPION: {
    text: "Entropion",
    rank: 501
    },

    FULL_ANAL_GLANDS: {
    text: "Full anal glands",
    rank: 702
    },

    GLAUCOMA: {
    text: "Glaucoma",
    rank: 200
    },

    HEART_MURMUR: {
    text: "Heart murmur",
    rank: 3
    },

    HEARTWORMS: {
    text: "Heartworms",
    rank: 2
    },
    
    HYPERTENSION: {
    text: "Hypertension (high blood pressure)",
    rank: 230
    },

    HYPERADRENOCORTICISM_PRESUMED: {
    text: "Hyperadrenocorticism (presumed)",
    rank: 201
    },

    HYPERADRENOCORTICISM: {
    text: "Hyperadrenocorticism",
    rank: 201
    },

    HYPOTHYROIDISM: {
    text: "Hypothyroidism",
    rank: 202
    },

    INFECTED_ANAL_GLANDS: {
    text: "Infected anal glands",
    rank: 204
    },

    KERATOCONJUNCTIVITIS_SICCA: {
    text: "Keratoconjunctivitis sicca (dry eye)",
    rank: 240
    },

    LEFT_SIDED_CONGESTIVE_HEART_FAILURE: {
    text: "Left sided congestive heart failure",
    rank: 1
    },

    LARYNGEAL_PARALYSIS: {
    text: "Laryngeal paralysis",
    rank: 401
    },

    MEIBOMIAN_GLAND_ADENOMA_PRESUMED: {
    text: "Meibomian gland adenoma (presumed)",
    rank: 310
    },

    MILD_PERIODONTAL_DISEASE: {
    text: "Mild periodontal disease",
    rank: 300
    },

    MODERATE_PERIODONTAL_DISEASE: {
    text: "Moderate periodontal disease",
    rank: 300
    },

    MYXOMATOUS_MITRAL_VALVE_DISEASE: {
    text: "Myxomatous mitral valve disease",
    rank: 40
    },

    NUCLEAR_SCLEROSIS: {
    text: "Nuclear sclerosis",
    rank: 970
    },

    OSTEOARTHRITIS: {
    text: "Osteoarthritis (arthritis)",
    rank: 100
    },

    OSTEOARTHRITIS: {
    text: "Osteoarthritis (arthritis)",
    rank: 100
    },

    OTITIS: {
    text: "Otitis externa (ear infection)",
    rank: 310
    },

    OVERWEIGHT: {
    text: "Overweight",
    rank: 700
    },

    PANCREATITIS: {
    text: "Pancreatitis",
    rank: 203
    },
    
    PARTIALLY_BLIND: {
    text: "Partially blind",
    rank: 900
    },

    PARTIALLY_VACCINATED: {
    text: "Partially vaccinated",
    rank: 980
    },

    PERIODONTAL_DISEASE: {
    text: "Periodontal disease",
    rank: 300
    },

    REVERSE_SNEEZING: {
    text: "Reverse sneezing",
    rank: 800
    },

    SECOND_DEGREE_AV: {
    text: "2nd Degree Atrioventricular Block",
    rank: 610
    },

    SEVERE_PERIODONTAL_DISEASE: {
    text: "Severe periodontal disease",
    rank: 300
    },

    UNDERWEIGHT: {
    text: "Underweight",
    rank: 820
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

      // FIX 1: Lenient Header Detection
      // We check for Bold OR Underline at the start of the paragraph.
      // This ensures "Weight:" (Bold only) is recognized as a header.
      function isHeaderParagraph(p) {
      const text = p.getText();
      if (!text.trim().length) return false;
      const textObj = p.editAsText();
      return textObj.isBold(0) || textObj.isUnderline(0);
      }

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
      if (nextChild.getType() === DocumentApp.ElementType.PARAGRAPH && nextChild.asParagraph().getText().trim() === "") {
      try { nextChild.removeFromParent(); } catch (e) { break; }
      } else { break; }
      }
      }

      const targetBase = summaryIndex === -1 ? body.getNumChildren() : summaryIndex + 1;

      templateBuffer.forEach((newTmpl) => {
      let inserted = false;
      const newRank = newTmpl.rank || 999;

      for (let i = targetBase; i < body.getNumChildren(); i++) {
      const child = body.getChild(i);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

      const p = child.asParagraph();
      const pText = p.getText().replace(/\u00A0/g, " ").trim().toLowerCase();

      // FIX 2: Continue, don't Break
      // If the line is empty, skip it. If it's not a header, skip it.
      // Do NOT 'break', otherwise we stop looking before we hit the Wellness header.
      if (!pText) continue; 
      if (!isHeaderParagraph(p) || !pText.includes(":")) continue;

      const headerText = pText.split(":")[0];
      let existingRank = null;

      // FIX 3: Robust Fuzzy Matching
      // We look for any overlap between the header (e.g. "Weight") 
      // and the registry (e.g. "Overweight")
      for (const key in allDiags) {
      const diagText = allDiags[key].text.toLowerCase();

      // Match if the header is inside the diagnosis (Weight is in Overweight)
      // OR the diagnosis is inside the header (Overweight contains Weight)
      if (diagText.includes(headerText) || headerText.includes(diagText)) {
      existingRank = allDiags[key].rank;
      break;
      }
      }

      // Default rank for recognized headers not in registry is 999 (like Wellness)
      let effectiveRank = (existingRank !== null) ? existingRank : 999;

      // If the header in the doc is lower priority (higher rank) than our new one, 
      // insert the new one right here.
      if (effectiveRank > newRank) {
      insertTemplateAtIndex(body, newTmpl, i);
      inserted = true;
      break;
      }
      }

      // --- FALLBACK ---
      if (!inserted) {
      let fallbackIdx = targetBase;
      while (fallbackIdx < body.getNumChildren()) {
      const next = body.getChild(fallbackIdx);
      // Put it at the first empty line or the end of the block
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

  // Vaccines Registry
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

  // Labwork Registry
    HEARTWORMS_PREVENTION_HEADER: 
    'Heartworms prevention:',

    HEARTWORM_TEST: g => 
    `A heartworm test was performed on your ${g.dog}. We will contact you in 3 - 4 business days with the results.`,

    LAB_RESULTS: g => 
    `Samples were drawn from your ${g.dog}. You will receive a call in 3 - 4 business days with the results.`,
    
    LABWORK: 
    'Early detection labwork:',

    HW_PREVENTION_SENTENCE: g =>
    `Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your ${g.dog} on monthly preventatives such as Heartgard, Simparica Trio, Revolution, etc.`,

  // Spay & Neuter Registry
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

  // Diet Registry
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

  // Dental Registry
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

  // Weight Management Registry
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

  // Ophthalmology Registry
    BLIND_HEADER:
    "Blind:",

    BLIND_OBSTACLE_COURSE: 
    /Make sure to keep your living space free of obstacles to prevent your (dog|cat) from tripping or bumping into things by mistake/i,

    CHERRY_EYE_HEADER: 
    /Cherry eye(s)?:/i,

    CHERRY_EYE_SURGERY_RECOMMENDATION: 
    /For dogs older than 1 year it is best to have the cherry eye(s)? corrected as soon as possible/i,

    CHERRY_EYE_IN_DOGS_AND_CATS_ARTICLE: { 
      text: "Cherry Eye in Dogs and Cats", 
      url: "https://vcahealthcare.com/know-your-pet/cherry-eye-in-dogs" 
    },

    COMPLETE_CATARACT_HEADER:
    /Complete cataract(s)?:/i,

    CONJUNCTIVITIS_HEADER:
    "Conjunctivitis:",

     CORNEAL_ULCER_HEADER: 
    /Corneal ulcer(s)?:/i,

    ENTROPION_HEADER:
    "Entropion:",

    EYE_DROP_MEDS:
    "Apply eye drops before eye ointments & wait 5 minutes between all eye medications to allow time for absorption.",

    GLAUCOMA_HEADER:
    "Glaucoma:",

    GLAUCOMA_RECHECK:
    "These must be given every 8 hours consistently, and many dogs need this lifelong. Bring your dog back in 1 week for a recheck of the eye pressure. If you notice worsening redness of the eyes, pain, discomfort when touching the head, or your dog squinting more, contact the clinic immediately.",

    HALO_HARNESS_ARTICLE: {
      text: "Halo harness",
      url: "https://www.muffinshalo.com/"
    },

    KCS_ARTICLE : {
      text: "Dry Eye (Keratoconjunctivitis Sicca) in Dogs and Cats article",
      url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951823`
    },

    KCS_HEADER:
    "Keratoconjunctivitis sicca:",

    KCS_MEDS_INCREASE:
    "If you notice green mucus around the eye, redness, or excessive scratching, it is possible your dog requires stronger medicine to continue controlling the eye.",

    KCS_SCHEDULE:
    "Schedule a recheck in 6 weeks so we can double check if the medicine needs to be increased.",

    KCS_TIMELINE:
    "This is a lifelong disease that, similar to allergies, is managed but not cured.",

    LENTICULAR_SCLEROSIS_IN_DOGS_ARTICLE : {
    text: "Lenticular Sclerosis in Dogs",
    url: `https://vcahospitals.com/know-your-pet/lenticular-sclerosis-in-dogs`
    },

    LUBRICATE_MEIBOMIAN_GLAND_ADENOMA:
    "Use a lubricating eye drop twice daily until the mass is removed to prevent damage to the eye.",

    MEIBOMIAN_GLAND_ADENOMA_HEADER:
    "Meibomian gland adenoma:",

    NUCLEAR_SCLEROSIS_HEADER:
    "Nuclear sclerosis:",

    REMOVE_MEIBOMIAN_GLAND_ADENOMA:
    "It is recommended to get this removed as soon as possible in healthy dogs to prevent secondary ulcer formation.",

    MEIBOMIAN_GLAND_ADENOMA_ARTICLE: {
      text: "Meibomian Gland (Eyelid) Tumors in Dogs article",
      url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&catId=254095&id=10194756`
    },

    TRIPPING_HAZARD:
    "Make sure to keep your living space free of obstacles to prevent your dog from tripping or bumping into things by mistake.",
  
  // Cardiology Registry
    THE_AMERICAN_HEARTWORM_SOCIETY_STATEMENT_ARTICLE : {
    text: "the American Heartworm Society statement",
    url: `https://www.heartwormsociety.org/resources/65-clinical-faqs/507-the-ahs-protocol-vs-slow-kill`
    },

    CANINE_HEARTWORM_GUIDELINES_BY_THE_AMERICAN_HEARTWORM_SOCIETY_ARTICLE : {
    text: "Canine Heartworm Guidelines by the American Heartworm Society.",
    url: `https://www.heartwormsociety.org/veterinary-resources/american-heartworm-society-guidelines`
    },

    CANINE_HEARTWORMS_AND_PREVENTING_DISEASE_ARTICLE : {
    text: "Canine Heartworms and Preventing Disease",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=11942142`
    },
    
    CONGESTIVE_HEART_FAILURE_IN_DOGS_CATS_ARTICLE_ARTICLE : {
    text: "Congestive Heart Failure in Dogs & Cats article",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=8501760`
    },

    CONTINUE_HEARTWORM_PREVENTION:
    "heartworm prevention is given for another month",

    DOXYCYCLINE_FOUR_WEEKS:
    "Doxycycline is given for 4 weeks.",

    EKG_RECOMMENDATION:
    "An EKG should be performed every 6 months to ensure there are no changes to your dog’s heart rhythm.",
    
    FIRST_MONTH_HEADER:
    "1st Month:",

    FOURTH_MONTH_HEADER:
    "4th Month:",
    
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

    HEART_MURMUR_WARNING_SIGNS:
    `If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease. Contact the clinic immediately.`,

    HEARTWORM_DISEASE_HEADER:
    "Heartworms:",

    HEARTWORMS_CAGE_REST:
    "It is vital you cage rest your dog throughout the ENTIRE treatment course & for 2 weeks after the last injection. Failure to do so can cause clots to form in the heart & blood vessels which can be fatal.",

    HEARTWORMS_EMERGENCY:
    "Excessive sluggishness, respiratory distress, and coughing up blood are signs of an emergency.",

    HEARTWORM_PREVENTION_COMPARISON_CHART_FOR_DOGS_AND_CATS_ARTICLE : {
    text: "Heartworm Prevention Comparison Chart for Dogs and Cats",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951375`
    },

    HEARTWORMS_PREVENTION_GIVEN:
    "Heartworm prevention is also given",

    HEARTWORMS_PREVENTION_HEADER:
    "Heartworms prevention:",

    HEARTWORMS_SLOW_KILL_WARNING:
    "Note that it can take years for heartworms to be completely eradicated via the slow kill method. There is still a chance of death during this time.",

    HEARTWORM_TEST_REPEAT:
    "For that reason it’s important your dog comes back in 6 months for a repeat of the heartworm test.",

    HIGH_BLOOD_PRESSURE_IN_OUR_PETS_ARTICLE : {
    text: "High Blood Pressure in Our Pets",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951756`
    },

    HYPERTENSION_DIAGNOSIS:
    "Your dog has elevated blood pressure.",

    HYPERTENSION_HEADER:
    "Hypertension",

    HYPERTENSION_MEDICINE:
    "This is a lifelong disease which can be managed but not cured. Your dog will PERMANENTLY require medication.",

    LEFT_SIDED_CONGESTIVE_HEART_FAILURE_HEADER:
    "Left sided congestive heart failure:",

    MELARSOMINE_INJECTION_SCHEDULE:
    "Medication to kill adult heartworms is injected into the back.",

    MITRAL_VALVE_DISEASE_IN_DOGS_ARTICLE : {
    text: "Mitral Valve Disease in Dogs",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=8526511`
    },

    MYXOMATOUS_MITRAL_VALVE_DISEASE_HEADER:
    "Myxomatous mitral valve disease:",

    PREDNISONE_SCHEDULE:
    "Prednisone (a steroid) is given for a month to control inflammation",

    PREDNISONE_TAPERING:
    "The medicine should be tapered as described below.",

    PREVENTING_HEARTWORM_INFECTION_IN_DOGS_ARTICLE : {
    text: "Preventing Heartworm Infection in Dogs",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&catId=102894&id=4951473`
    },

    SCHEDULE_ECHO: g =>
    `an echocardiogram to look at the inner workings of the heart and diagnose the cause of the disease will need to be scheduled.`,

    SCHEDULE_ECHOS: g =>
    `An echocardiogram to look at the inner workings of the heart and diagnose the cause of the disease will need to be scheduled.`,

    SECOND_DEGREE_ATRIOVENTRICULAR_BLOCK_HEADER:
    "2nd Degree Atrioventricular Block",

    SECOND_MONTH_HEADER:
    "2nd Month:",

    THIRD_MONTH_HEADER:
    "3rd Month:",

  // Respiratory Registry
    BOAS_GENETICS:
    `Your dog’s breed is known to have a disorder called brachycephalic obstructive airway syndrome.`,

    BOAS_COLD_WATER_SHOCK: g =>
    `DO NOT pour cold water or ice on ${g.him} or the towel as this can cause the blood vessels to constrict, thereby putting your ${g.dog} in a state of shock.`,

    BORDETELLOSIS_HEADER:
    "Bordetellosis:",

    BORDETELLA_DISEASE:
    "Bordetella",

    BRACHYCEPHALIC_OBSTRUCTIVE_AIRWAY_SYNDROME_HEADER:
    `Brachycephalic obstructive airway syndrome:`,

    CHRONIC_BRONCHITIS_HEADER:
    "Chronic bronchitis: ",

    CHRONIC_BRONCHITIS_IN_DOGS_ARTICLE_ARTICLE : {
    text: "Chronic Bronchitis in Dogs article",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&catId=102899&id=5138343`
    },

    COLLAPSING_TRACHEA_HEADER: "Collapsing trachea:",
    TRACHEAL_COLLAPSE_IN_DOGS_ARTICLE_ARTICLE: {
    text: "Tracheal Collapse in Dogs article",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951968`
    },

    COLLAPSING_TRACHEA_NO_TREATMENT:
    "At this time no medication is being used as your dog appears to be well controlled at home without it.",

    ELONGATED_SOFT_PALATE_TX:
    `Treatment is using surgery to remove the elongated soft palate.`,

    LARYNGEAL_PARALYSIS_HEADER:
    "Laryngeal paralysis:",
    
    LARYNGEAL_PARALYSIS_IN_DOGS_ARTICLE_ARTICLE: {
    text: "Laryngeal Paralysis in Dogs article",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&catId=102899&id=4952489`
    },

    REVERSE_SNEEZING_HEADER:
    "Reverse sneezing:",

    REVERSE_SNEEZING_TREATMENT:
    " includes massaging the throat when it occurs or and giving Benadryl 25mg (give 1 tablet per 25 lbs every 12 hours) or Zyrtec 10mg (give up to 1 tablet per 10 lbs every 12 hours) for allergies.",

    STENOTIC_NARES_TX:
    `Treatment is surgery to open up the nares more.`,
    
    OVERHEATING_HEADER:
    `Overheating:`,

  // Endocrine Registry
    CONTINUE_MONITOR_DIABETES:
    "Continue to monitor for signs of uncontrolled diabetes such as increased thirst, urination, appetite, weight loss, and seizures.",
    
    CUSHINGS_SYNDROME_HYPERADRENOCORTICISM_ARTICLE: {
    text: "Cushing's Syndrome (Hyperadrenocorticism) article",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951495`
    },

    DIABETES_EMERGENCY:
    "If signs persist, seek immediate medical treatment.",

    DIABETES_HEADER:
    "Diabetes mellitus:",
    
    DIABETES_MELLITUS_INTRODUCTION_ARTICLE: {
    text: "Diabetes Mellitus Introduction",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951506`
    },

    DIABETES_SYMPTOMS:
    " of diabetes include increased thirst, urination, appetite, weight loss, and seizures.",

    GLUCOSE_CURVE_ADVISED:
    "Your dog will need to come back in two weeks for a glucose curve.",

    HYPERADRENOCORTICISM_HEADER:
    "Hyperadrenocorticism:",

    HYPOTHYROIDISM_DIAGNOSED:
    "Your dog has been diagnosed with hypothyroidism.",

    HYPOTHYROIDISM_HEADER:
    "Hypothyroidism:",

    HYPOTHYROIDISM_IN_DOGS_ARTICLE_ARTICLE: {
    text: "Hypothyroidism in Dogs article",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4952004`
    },

    HYPOTHYROID_RECHECK:
    "We will need to recheck thyroid hormone levels in 4 - 6 weeks to ensure we’re not giving too much thyroid & causing hyperthyroidism. Give the medicine 4 - 6 hours BEFORE the appointment.",

    HYPERADRENOCORTICISM_REPEAT_TEST:
    "Continue to monitor for side effects including vomiting, diarrhea, listlessness, or decreased water intake. A repeat test is necessary 7 days after starting the new medication.",

    HYPERADRENOCORTICISM_SYMPTOMS:
    "Continue to monitor for side effects including vomiting, diarrhea, listlessness, or decreased water intake.",

    INSULIN_ADMINISTRATION_WARNING:
    "Give insulin AFTER your pet has eaten. Do not mix up 100 unit insulin syringes and 40 unit insulin syringes.",

    LIGHT_KARO_SYRUP_ARTICLE: {
    text: "light Karo syrup",
    url: `https://a.co/d/07pkwuVd`
    },

    PANCREATITIS_HEADER:
    "Pancreatitis:",
    
    PANCREATITIS_IN_DOGS_ARTICLE: {
    text: "Pancreatitis in Dogs",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4952412`
    },

    PANCREATITIS_PREVENTION:
    "The best way to prevent pancreatic flare ups is by feeding a low fat diet and cutting out treats that are high in fat.",

    RECHECK_ACTH_STIM_TEST:
    "If trilostane is given, a recheck ACTH stim test should be performed two weeks after starting meds to make sure your dog is receiving the right amount.",

    TRILOSTANE_HYPOADRENOCORTICISM:
    "However, it’s important we don’t cause an underproduction of cortisol.",

    TRILOSTANE_WARNING:
    "Make sure to give the trilostane 4 - 6 hours before bringing your dog into the clinic. If your dog is feeling lethargic, vomiting, or has a decreased appetite, contact the clinic immediately and we can discuss having the test done sooner.",

  // Gastrointestinal Registry
    ACUTE_GASTROENTERITIS_HEADER:
    "Acute gastroenteritis:",

    ANAL_GLANDS_BRING_STOOL:
    "If your dog continues to fixate on the anus, bring back a stool sample & we can test for intestinal parasites.",

    ANAL_GLANDS_EXPRESSED:
    "Your dog had full anal glands that were expressed at the clinic.",

    ANAL_GLANDS_HEADER:
    "Anal glands:",

    ANAL_GLANDS_STILL_SCOOTING:
    "If your dog continues to scoot on the floor, bring back a stool sample & we can test for intestinal parasites.",
    
    ANAL_GLAND_SYMPTOMS:
    "If your dog’s anal glands are full, you may see scooting on the floor or over fixation on the anus.",

    EMPTYING_A_DOG_OR_CAT_S_ANAL_SACS_ARTICLE: {
    text: "Emptying a Dog or Cat's Anal Sacs",
    url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951501`
    },

    INFECTED_ANAL_GLANDS_SEEN:
    "Your dog had full anal glands that appeared infected when they were expressed in the clinic.",

    INFECTED_ANAL_GLANDS_HEADER: "Infected anal glands:",

    VOMITING_POST_MAROPITANT:
    "If you still see vomiting within 24 hours of the injection, your dog needs to go to your nearest veterinary emergency hospital immediately.",

  // Musculoskeletal Registry
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

  // Dermatology Registry
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

  // Immunology Registry
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
    `THERE IS NO CURE FOR RABIES. The only way to test for rabies involves decapitating an animal & taking samples of the brain. State law requires any animal that is exposed to rabies to either undergo quarantine for up to 6 months or be euthanized.`,

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

  // General Illness Registry
    COMMON_CAUSES:
    /Common causes/i,

    CONTINUE_MEDICATION_AS_PRESCRIBED:
    "Continue your current medication as previously prescribed",
    
    CONTINUE_MEDICATION_LABWORK_6_MONTHS:
    "Continue to give the medication as you have been. Bloodwork is recommended every 6 months.",

    DIAGNOSE:
    /Diagnose/i,
    
    DIAGNOSIS:
    /Diagnosis/i,
    
    E_COLLAR_ADVISE:
    "Keep a hard e collar on for the next 14 days.",

    E_COLLAR_HEADER:
    "E collar:",

    E_COLLAR_MONITOR:
    "but you must monitor your dog all throughout & replace the e collar immediately.",

    RECHECK_ADVISE_1_WEEK:
    "Bring your dog back in 1 week for a recheck appointment if no improvement is seen (return immediately if worsening).",

    RECHECK_ADVISE_3_DAYS:
    "Bring your dog back in 3 days for a recheck appointment if no improvement is seen (return immediately if worsening).",
    
    SYMPTOMS:
    /Symptoms/i,

    TREATMENT:
    /Treatment/i,

    YOU_WILL_BE_CALLED_WITH_RESULTS:
    "You will be called with results in 3 - 4 business days.",
    };

/* ------------------ Reset & Replacement ------------------ */
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
    function reverseGenerateTemplate(sex, plurality) {
    const body = DocumentApp.getActiveDocument().getBody();
    const reverseMap = getReverseRegistryMap(); 

    // Discovery Buffers
    globalThis.globalNewFormatEntries = {};
    const NEW_MEDS_FOUND = {}; 
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

    // 1. Process Paragraphs for the template text
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
    const paragraph = element.asParagraph();
    let text = paragraph.getText() || "";
    if (!text.trim() || text.includes("/* --- REVERSE")) continue;

    paragraphs.push("`" + escapeBackticks(convertToGPronouns(text)) + "`");

    extractFormattingSpans(paragraph.editAsText(), {
    boldKeys, boldUnderlineKeys, italicKeys, greenKeys, redKeys, linkKeys
    }, reverseMap, NEW_LINKS_REGISTRY);
    }

    // 2. Process Tables to discover NEW medications
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

    // Check global MEDICINE_REGISTRY for existing entries
    const alreadyExists = MEDICINE_REGISTRY[drugName] || 
    Object.values(MEDICINE_REGISTRY).some(m => m.label.toLowerCase() === label.toLowerCase());

    if (!alreadyExists) {
    let cleanInstructions = stripMedPrefix(rowData[1] || "");
    const newEntry = {
    label: label, 
    instructions: cleanInstructions,
    class: rowData[2] || "Unknown", 
    sideEffects: rowData[3] || "Unknown"
    };

    // Track locally so it prints at the end, and update main registry for this session
    MEDICINE_REGISTRY[drugName] = newEntry;
    NEW_MEDS_FOUND[drugName] = newEntry;
    }
    }
    }
    }

    const formatKeys = (set) => Array.from(set).sort().map(k => `"${k}"`).join(",\n      ");

  // Generated template
    let outputCode = `// Generated Template
    function generateTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: [""],
    text: [
    ${paragraphs.join(",\n      ")}
    ].join('\\n'),\n\n`

    if (boldKeys.size > 0) outputCode += `    boldKeys: [\n      ${formatKeys(boldKeys)}\n    ],\n\n`;
    if (boldUnderlineKeys.size > 0) outputCode += `    boldUnderlineKeys: [\n      ${formatKeys(boldUnderlineKeys)}\n    ],\n\n`;
    if (italicKeys.size > 0) outputCode += `    italicKeys: [\n      ${formatKeys(italicKeys)}\n    ],\n\n`;
    if (greenKeys.size > 0) outputCode += `    greenKeys: [\n      ${formatKeys(greenKeys)}\n    ],\n\n`;
    if (redKeys.size > 0) outputCode += `    redKeys: [\n      ${formatKeys(redKeys)}\n    ],\n\n`;
    if (linkKeys.size > 0) outputCode += `    linkKeys: [\n      ${formatKeys(linkKeys)}\n    ],\n\n`;

    outputCode += `  };\n}\n`;

    // Append New Format Registry Entries (discovered during this run)
    const newFormatKeys = Object.keys(globalNewFormatEntries);
    const newLinkKeys = Object.keys(NEW_LINKS_REGISTRY);

    if (newFormatKeys.length > 0 || newLinkKeys.length > 0) {
    outputCode += "\n/* --- NEW FORMAT REGISTRY ENTRIES --- */\n";
    newFormatKeys.sort().forEach(key => {
    outputCode += `${key}: "${globalNewFormatEntries[key]}",\n`;
    });
    newLinkKeys.sort().forEach(key => {
    const link = NEW_LINKS_REGISTRY[key];
    outputCode += `${key}: {\n  text: "${link.text.replace(/"/g, '\\"')}",\n  url: \`${link.url}\`\n},\n`;
    });
    }

    // Append ONLY NEW Medicine Registry entries
    const newMedKeys = Object.keys(NEW_MEDS_FOUND);
    if (newMedKeys.length > 0) {
    outputCode += "\n/* --- NEW MEDICINE REGISTRY ENTRIES --- */\n";
    newMedKeys.sort().forEach(key => {
    const med = NEW_MEDS_FOUND[key];
    outputCode += `  ${key}: {\n    label: "${med.label.replace(/"/g, '\\"')}",\n    instructions: "${med.instructions.replace(/"/g, '\\"')}",\n    class: "${med.class}",\n    sideEffects: "${med.sideEffects}"\n  },\n`;
    });
    }

    body.appendPageBreak();
    body.appendParagraph("/* --- REVERSE GENERATED CODE --- */").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(outputCode);
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
    } 
    else if (val instanceof RegExp) {
    // Strips regex slashes and flags to get the searchable text
    const regexSource = val.source.replace(/\\/g, "").toLowerCase();
    reverseMap.textToKey[regexSource] = key;
    } 
    else if (val && typeof val === 'object') {
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
    const isBold = textElement.isBold(start), 
    isUnderline = textElement.isUnderline(start),
    isItalic = textElement.isItalic(start), 
    color = textElement.getForegroundColor(start),
    linkUrl = textElement.getLinkUrl(start);

    let end = start;
    while (end < text.length && textElement.isBold(end) === isBold && textElement.isUnderline(end) === isUnderline &&
    textElement.isItalic(end) === isItalic && textElement.getForegroundColor(end) === color && 
    textElement.getLinkUrl(end) === linkUrl) {
    end++;
    }

    let rawSpan = text.substring(start, end);
    let spanClean = cleanSpaces(rawSpan);

    if (spanClean.length > 1) {
    let existingKey = reverseMap.textToKey[spanClean.toLowerCase()] || (linkUrl ? reverseMap.urlToKey[linkUrl] : null);

    if (linkUrl) {
    let linkKey = existingKey;
    if (!linkKey) {
    linkKey = spanClean.replace(/[^A-Za-z0-9]/g, "_").toUpperCase() + "_ARTICLE";
    newLinksRegistry[linkKey] = { text: spanClean, url: linkUrl };
    }
    keys.linkKeys.add(linkKey);
    } else {
    const isGreen = (color === "#6aa84f" || color === "#008000" || color === "#b6d7a8"), 
    isRed = (color === "#ff0000" || color === "#ea9999");

    let identifier = existingKey;

    // Auto-generate Header keys if not in registry
    if (!identifier && isBold && spanClean.endsWith(':')) {
    let baseName = spanClean.replace(/[:]/g, "").replace(/[^A-Za-z0-9]/g, "_").toUpperCase();
    identifier = baseName.replace(/_+/g, "_") + "_HEADER";
    globalNewFormatEntries[identifier] = spanClean;
    }

    if (!identifier) identifier = spanClean;

    if (isGreen) {
    keys.greenKeys.add(identifier);
    } else if (isRed) {
    keys.redKeys.add(identifier);
    } else {
    if (isBold && isUnderline) keys.boldUnderlineKeys.add(identifier);
    else if (isBold) keys.boldKeys.add(identifier);
    if (isItalic) keys.italicKeys.add(identifier);
    }
    }
    }
    start = end;
    }
    }

    function stripMedPrefix(text) {
    if (!text) return "";
    let cleanText = text.trim();
    if (typeof MED_PREFIX === 'undefined') return cleanText;

    for (let key in MED_PREFIX) {
    let prefixText = MED_PREFIX[key];
    if (cleanText.startsWith(prefixText)) {
    cleanText = cleanText.substring(prefixText.length).replace(/^[\n\r\s]+/, "");
    break; 
    }
    }
    return cleanText;
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
    `Heartworms prevention: Heartworms are spread by mosquitoes which don’t die in the Texas "winter", so our pets are at risk of infection year round. Furthermore, heartworms can be fatal & there is a risk of death even with proper treatment. Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your ${g.dog} on monthly preventatives such as Heartgard, Simparica Trio, Revolution, etc. Depending on the brand, they can protect your ${g.dog} from heartworms, fleas, ticks, & common intestinal parasites with a single treatment. These can be given orally or topically & are generally well tolerated. Because your ${g.dog} ${g.is} still growing, you will need to come back once a month to have ${g.him} weighed & get the appropriate dose of preventative.`,
    `${spayorneuterText}`,
    `Food: A high quality diet is the best way to keep your ${g.dog} healthy. Any puppy diet from Hill’s Science Diet (Hill's puppy dry food or Hill's puppy wet food), Purina Pro Plan (Purina puppy dry food or Purina puppy wet food), or Royal Canin (RC puppy dry food or RC puppy wet food) are all acceptable. A puppy diet is advised until your ${g.dog} ${g.is} a year old at which point you can transition to an adult diet. Dry food & wet food are both appropriate to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag or can for a dog of ${g.his} weight.`,
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
    'HEARTWORMS_PREVENTION_HEADER',
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

    const text = [
        `Vaccines: Your ${g.dog} ${g.has} received ${g.his} first ${g.round} of adult vaccinations. The 1 year rabies vaccine was given in the right hindlimb. The initial distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo ${g.shot} with the initial lepto vaccine in the left hindlimb. The 1 year bordetella vaccine was given orally. Your ${g.dog} will need a booster of the DAPP and lepto vaccines in 3 - 4 weeks.`,
        
        `You may notice that after vaccination your ${g.dog} ${g.is} more tired than usual, ${g.eats} less, or ${g.is} sore at the injection sites, & this is perfectly normal. Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your ${g.dog} back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your ${g.dog}.`,
        
        `Heartworms prevention: A heartworm test was performed on your ${g.dog}. We will contact you in 3 - 4 business days with the results. Heartworms are spread by mosquitoes which don’t die in the Texas "winter", so our pets are at risk of infection year round. Furthermore, heartworms can be fatal & there is a risk of death even with proper treatment. Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your ${g.dog} on monthly preventatives such as Heartgard, Nexgard, Simparica Trio, Revolution, etc.`,
        
        `Early detection labwork: Yearly blood work is recommended for ${g.dogs} the same as it is in humans and starts at 3 years of age. This lets us get a baseline for your pet and allows us to catch abnormalities before they’re noticeable outwardly. Depending on the panel run, this can check for issues in the liver, kidneys, thyroid, bladder, glucose, and many other organs and values. At 6 years of age, a larger panel for “senior” pets is advised.`,
        
        `Food: A high quality diet is the best way to keep your ${g.dog} healthy. Food from Hill’s Science Diet (Hill's dog dry food or Hill's dog wet food), Purina Pro Plan (Purina dog dry food or Purina dog wet food), or Royal Canin (RC dog dry food or RC dog wet food) are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in ${g.dogs}, so either is wonderful to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag or can for a dog of ${g.his} weight.`,

        `Dental care: The best way to keep your ${g.dogs} teeth healthy is to brush them daily for 10 seconds total using a small dog toothbrush, medium/large dog toothbrush, & animal safe toothpaste. Animal safe toothpaste such as C.E.T. can be purchased from the clinic or from online stores. Getting your ${g.dog} used to having ${g.his} teeth brushed early will improve ${g.his} overall health.`,
        
        `You can start by having ${g.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then applying the pet safe toothpaste & letting ${g.him} lick it off every day for a week. Finally, gently brush ${g.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
        
        `If your ${g.dog} resists having ${g.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${g.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`,
        
        `Next appointment: Bring your ${g.dog} back in 3 - 4 weeks for a booster of ${g.his} vaccines.`
    ].join('\n');

    return {
    sex,
    plurality,
    text,
    boldKeys: [
    'VACCINES_HEADER',
    'HEARTWORMS_PREVENTION_HEADER',
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

    // 3. Main Template Text
      const text = [ 
      `Vaccines: Your ${g.dog} ${g.has} received ${g.his} first ${g.round} of adult vaccinations. Because you have kept to ${g.his} vaccination schedule, ${g.his} immune system will not need another booster until next year.`,
      `The 1 year rabies vaccine was given in the right hindlimb. The 1 year distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo ${g.shot} with the 1 year lepto vaccine in the left hindlimb. The 1 year bordetella vaccine was given orally. You may notice that after vaccination your ${g.dog} ${g.is} more tired than usual, ${g.eats} less, or ${g.is} sore at the injection sites, & this is perfectly normal.`,
      `Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your ${g.dog} back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your ${g.dog}.`,
      `Heartworms prevention: A heartworm test was performed on your ${g.dog}. We will contact you in 3 - 4 business days with the results. Heartworms are spread by mosquitoes which don’t die in the Texas "winter", so our pets are at risk of infection year round. Furthermore, heartworms can be fatal & there is a risk of death even with proper treatment. Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your ${g.dog} on monthly preventatives such as Heartgard, Nexgard, Simparica Trio, Revolution, etc.`,
      `Early detection labwork: Samples were drawn from your ${g.dog}. You will receive a call in 3 - 4 business days with the results. Yearly blood work is recommended for dogs the same as it is in humans for the sake of monitoring for abnormalities that aren’t visible from the outside. Depending on the panel run, this can check for issues in the liver, kidneys, thyroid, bladder, glucose, and many other organs and values. If no abnormalities are found, the results can be used as a baseline so that your ${g.dogs} overall health is closely monitored.`,
      `Food: A high quality diet is the best way to keep your ${g.dog} healthy. If you haven’t already, you can transition ${g.him} from ${g.his} ${g.puppy} diet to ${g.his} adult diet. Food from Hill’s Science Diet (Hill's dog dry food or Hill's dog wet food), Purina Pro Plan (Purina dog dry food or Purina dog wet food), or Royal Canin (RC dog dry food or RC dog wet food) are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in dogs, so either is wonderful to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag or can for a dog of ${g.his} weight.`,
      `Dental care: The best way to keep your ${g.dogs} teeth healthy is to brush them daily for 10 seconds total using a small dog toothbrush, medium/large dog toothbrush, & animal safe toothpaste. Animal safe toothpaste such as C.E.T. can be purchased from the clinic or from online stores. Getting your ${g.dog} used to having ${g.his} teeth brushed early will improve ${g.his} overall health.`,
      `You can start by having ${g.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then applying the pet safe toothpaste & letting ${g.him} lick it off every day for a week. Finally, gently brush ${g.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
      `If your ${g.dog} resists having ${g.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${g.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`,
      `Next appointment: Bring your ${g.dog} back one year from today for ${g.his} next annual vaccines.`
      ].join('\n');

    return {
    sex,
    plurality,
    text,
    diagnoses: [""],
    rank: 999,
    boldKeys: [
    'VACCINES_HEADER',
    'HEARTWORMS_PREVENTION_HEADER',
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
        `Food: A high quality diet is the best way to keep your ${g.dog} healthy. ${g.Dogs} that are older than 7 years are advised to be on a senior diet. Food from Hill’s Science Diet (Hill's senior dog dry food or Hill's senior dog wet food), Purina Pro Plan (Purina senior dog dry food or Purina senior dog wet food), or Royal Canin (RC senior dog dry food or RC senior dog wet food) are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in dogs, so either is wonderful to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag or can for a dog of ${g.his} weight.`
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
        `Food: A high quality diet is the best way to keep your ${g.dog} healthy. ${g.Dogs} that are older than 7 years are advised to be on a senior diet. Food from Hill’s Science Diet (Hill's senior dog dry food or Hill's senior dog wet food), Purina Pro Plan (Purina senior dog dry food or Purina senior dog wet food), or Royal Canin (RC senior dog dry food or RC senior dog wet food) are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in dogs, so either is wonderful to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag or can for a dog of ${g.his} weight.`
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
    plurality,
    text,
    diagnoses: ["OVERWEIGHT"],
    cleanupKeys: ["DIET_HEADER"],
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
    plurality,
    text,
    diagnoses: ["OVERWEIGHT"],
    cleanupKeys: ["DIET_HEADER"],
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
    function generateCanineHealthyWeightTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Weight: Your dog is a healthy weight for a dog of ${g.his} size. ${g.His} ribs can be felt without difficulty and ${g.he} has a slight waist. Keeping ${g.him} around ${g.his} current weight will help ${g.him} live approximately 1 ½ years longer than ${g.he} would if ${g.he} were over or underweight.`,
    `Continue to monitor ${g.his} weight and feed ${g.him} as you’ve been doing. Signs of an overweight dog include difficulty feeling the ribs and loss of a waist when viewed from above. You can switch ${g.his} treats to apple slices, carrots, green beans, ice cubes, or cucumbers if you notice ${g.him} starting to gain weight. Signs of an underweight dog are the spine being visible in the same fashion as your knuckles, ribs visible enough to be counted, and hips that can be felt when running your hand over your dog’s back end.`
    ].join('\n');

    return {
    sex,
    plurality,
    text,
    rank: 100,
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
    function generateCanineUnderweightTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Underweight: Your dog weighs less than the average dog of ${g.his} size. Ideally, we would be able to feel ${g.his} ribs but not see them. Helping ${g.him} gain weight can increase ${g.his} quality of life.`,
    `The best way for ${g.him} to gain weight is through diet. Food from Hill’s Science Diet, Purina Pro Plan, or Royal Canin are all wonderful diets as they’re formulated by veterinary scientists. You can also add lukewarm water to the food or low sodium chicken broth to increase the smell and flavor.`,
    `Increase how much ${g.he} eats by as much as 25 - 50%. We're aiming to have ${g.him} gain approximately 10% of ${g.his} current weight. Failure to gain weight is concerning for disease and would prompt us to perform tests such as labwork, ultrasound, or x-rays.`
    ].join('\n');

    return {
    sex,
    plurality,
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
    function generateCanineVaccineInformationTemplate(sex, plurality) {
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
    sex,
    plurality,
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
  // Blind | 0th, Partial
    function generateCanineBlind0PartialTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
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

  // Blind | 1st, Diagnosed
    function generateCanineBlind1Template(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Blind: Your dog shows signs of being completely blind. Make sure to keep your living space free of obstacles to prevent your dog from tripping or bumping into things by mistake. Avoid making changes to your living space as your dog has most likely memorized the layout and will be confused if things move around. You can also use a Halo harness or similar devices to prevent your pet from running into objects.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Blind | 2nd, Known
    function generateCanineBlind2KnownTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Blind: Your dog is known to be completely blind. Make sure to keep your living space free of obstacles to prevent your dog from tripping or bumping into things by mistake. Avoid making changes to your living space as your dog has most likely memorized the layout and will be confused if things move around. You can also use a Halo harness or similar devices to prevent your pet from running into objects.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Cherry Eye
    function generateCanineCherryEyeTemplate(sex, plurality = 'singular') {
    const g = getGrammar('eyes', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["CHERRY_EYE"],
    text: [
    `Cherry ${g.eye}: Your dog has ${g.cherry_eye}. This means that the ${g.gland} of the ${g.eye} that ${g.make} most of the tears ${g.is} poking out of ${g.its} normal position. Over time ${g.this} ${g.gland} can dry up & produce less tears, leading to a disorder known as dry eye. For dogs older than 1 year it is best to have the cherry ${g.eye} corrected as soon as possible. Surgery involves tucking the gland back in its normal position & using suture to prevent it from popping out again. You can learn more about cherry eyes from the Cherry Eye in Dogs and Cats article on Veterinary Partner.`
    ].join('\n'),

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

  // Complete Cataracts
    function generateCompleteCataractsTemplate(sex, plurality = 'singular') {
    const g = getGrammar('eyes', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["CATARACTS"],
    text: [
    `Complete ${g.cataract}: Your dog has ${g.a_cataract}. This typically forms due to age & completely blocks vision. While your dog may still see shadows & light, it is unlikely that very much vision is actually present in the ${g.eye}. Make sure to keep your living space free of obstacles to prevent your dog from tripping or bumping into things by mistake. Be mindful about moving silently as this may startle your dog.`
    ].join('\n'),
    boldKeys: [
    "COMPLETE_CATARACT_HEADER"
    ],

    boldUnderlineKeys: [
    "TRIPPING_HAZARD"
    ],
    };
    }

  // Conjunctivitis | Tests Declined
    function generateCanineConjunctivitisTestsDeclinedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["CONJUNCTIVITIS_PRESUMED"],
    text: [
    `Conjunctivitis: Your dog’s eyes may be inflamed due to keratoconjunctivitis sicca (also known as dry eye, checked by the Schirmer tear test) corneal ulcers (checked by the fluorescein eye stain) or glaucoma (checked by tonometry) among other diseases. At this time you’ve declined to perform these tests in favour of symptomatic treatment. An antibiotic/anti-inflammatory eye medication has been sent home. Use as directed below. Bring your dog back in 1 week for a recheck appointment if no improvement is seen (return immediately if worsening).`
    ].join('\n'),

    boldKeys: [
      "CONJUNCTIVITIS_HEADER"
    ],

    boldUnderlineKeys: [
      "RECHECK_ADVISE_1_WEEK"
    ],
    };
    }

  // Conjunctivitis | Diagnosed
    function generateCanineConjunctivitisDiagnosedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["CONJUNCTIVITIS_DIAGNOSED"],
    text: [
    `Conjunctivitis: Your dog’s eyes were checked for keratoconjunctivitis sicca (dry eye) via the Schirmer tear test, corneal ulcers via the fluorescein eye stain, & glaucoma via tonometry. At this time no signs of any of these diseases are present. As such, your dog has been diagnosed with conjunctivitis (inflammation of the eye due to irritation or infection). An antibiotic/anti-inflammatory eye medication has been sent home. Bring your dog back in 1 week for a recheck appointment if no improvement is seen (return immediately if worsening).`
    ].join('\n'),
    
    boldKeys: [
      "CONJUNCTIVITIS_HEADER"
    ],

    boldUnderlineKeys: [
      "RECHECK_ADVISE_1_WEEK"
    ],
    };
    }

  // Corneal Ulcer
    function generateCanineCornealUlcerTemplate(sex, plurality = 'singular') {
    const g = getGrammar('eyes', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["CORNEAL_ULCER"],
    text: [
    `Corneal ${g.ulcer}: ${g.a_corneal_ulcer} ${g.was} seen in your dog’s ${g.eye}. ${g.This} can occur due to excessive scratching, playing, or running into objects. An antibiotic has been sent home to resolve the inflammation & potential infection seen. Bring your dog back in 1 week for a recheck appointment if no improvement is seen (return immediately if worsening).`,
    `E collar: It’s important to keep a hard e collar on your dog to prevent rubbing the ${g.eye} on furniture or scratching ${g.it}. Keep a hard e collar on for the next 14 days. It needs to go at least an inch past your dog’s nose. Tighten the collar so that it can’t be pushed off, but make sure you can still get two fingers between the collar & the skin. It can be removed to allow for eating & when using the restroom but you must monitor your dog all throughout & replace the e collar immediately. Failure to keep the e collar may cause worsening infection or damage to occur.`
    ].join('\n'),
    boldKeys: [
    "CORNEAL_ULCER_HEADER",
    "E_COLLAR_HEADER"
    ],

    boldUnderlineKeys: [
    "E_COLLAR_ADVISE",
    "RECHECK_ADVISE_1_WEEK",
    "E_COLLAR_MONITOR"
    ],
    };
    }

  // Entropion
    function generateCanineEntropionTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["ENTROPION"],
    text: [
    `Entropion: Your dog’s eyelids turn too far towards the eyeballs. This causes the eyelashes to scratch against the eyes which causes inflammation, pain, and discomfort. Having entropion eyelids increases the risk of developing corneal ulcers. Surgery can be performed to correct this abnormality.`
    ].join('\n'),
    
    boldKeys: [
      "ENTROPION_HEADER"
    ],
    };
    }

  // Glaucoma | Diagnosed
    function generateCanineGlaucoma1DiagnosedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["GLAUCOMA"],
    text: [
    `Glaucoma: Normal eye pressure in a dog ranges from 10 - 25 mmHg, but your dog’s pressure measured significantly higher. The most common cause of glaucoma in dogs is genetics. To alleviate the pressure and the pain, we will start your dog on eye drops. These must be given every 8 hours consistently, and many dogs need this lifelong. Bring your dog back in 1 week for a recheck of the eye pressure. If you notice worsening redness of the eyes, pain, discomfort when touching the head, or your dog squinting more, contact the clinic immediately.`
    ].join('\n'),
    
    boldKeys: [
      "GLAUCOMA_HEADER"
    ],

    boldUnderlineKeys: [
      "GLAUCOMA_RECHECK"
    ],
    };
    }

  // Keratoconjunctivitis Sicca | 1st, Diagnosed
    function generateCanineKeratoconjunctivitisSicca1DiagnosedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["KERATOCONJUNCTIVITIS_SICCA"],
    text: [
    `Keratoconjunctivitis sicca: Your dog has been diagnosed with keratoconjunctivitis sicca (KCS), more commonly known as dry eye. This is an autoimmune disease where the immune system is attacking the part of the eye that produces tears. This is a lifelong disease that, similar to allergies, is managed but not cured.`,
    `Eye medication is used to suppress the immune system in the eye. Over time you may notice your dog needs stronger eye medication as the body becomes resistant to the initial medication. For now, apply the medication provided. You can also pick up non-medicated eye drops from any store & apply them twice a day for the next two weeks as the medicine takes time to activate. `,
    `Apply eye drops before eye ointments & wait 5 minutes between all eye medications to allow time for absorption. Schedule a recheck in 6 weeks so we can double check if the medicine needs to be increased. You can learn more from the Dry Eye (Keratoconjunctivitis Sicca) in Dogs and Cats article on Veterinary Partner.`
    ].join('\n'),
    boldKeys: [
      "KCS_HEADER"
    ],

    boldUnderlineKeys: [
      "EYE_DROP_MEDS",
      "KCS_SCHEDULE",
      "KCS_TIMELINE",
    ],

    linkKeys: [
      "KCS_ARTICLE"
    ],
    };
    }

  // Keratoconjunctivitis Sicca | 2nd, Controlled
    function generateCanineKeratoconjunctivitisSicca2ControlledTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["KERATOCONJUNCTIVITIS_SICCA"],
    text: [
    `Keratoconjunctivitis sicca: Your dog is known to have keratoconjunctivitis sicca (KCS), more commonly known as dry eye. This autoimmune disease is life long & requires continuous medication. Continue to give your dog’s eye medicine as previously prescribed. If you notice green mucus around the eye, redness, or excessive scratching, it is possible your dog requires stronger medicine to continue controlling the eye. You can learn more from the Dry Eye (Keratoconjunctivitis Sicca) in Dogs and Cats article on Veterinary Partner.`
    ].join('\n'),
    boldKeys: [
      "KCS_HEADER"
    ],

    boldUnderlineKeys: [
      "KCS_MEDS_INCREASE"
    ],

    linkKeys: [
      "KCS_ARTICLE"
    ],
    };
    }

  // Meibomian Gland Adenoma | 1st, Presumed
    function generateCanineMeibomianGlandAdenoma1PresumedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["MEIBOMIAN_GLAND_ADENOMA_PRESUMED"],
    text: [
    `Meibomian gland adenoma: The mass on your dog’s eyelid appears to be a meibomian gland adenoma. While these are typically benign, they can grow and cause obstruction of the vision  or scratch against the cornea & cause ulcers. It is recommended to get this removed as soon as possible in healthy dogs to prevent secondary ulcer formation. If the growth contacts the eye, you may see rubbing the face, redness in the eye, and mucoid discharge. Use a lubricating eye drop twice daily until the mass is removed to prevent damage to the eye. You can learn more by reading the Meibomian Gland (Eyelid) Tumors in Dogs article on Veterinary Partner.`
    ].join('\n'),
    boldKeys: [
      "MEIBOMIAN_GLAND_ADENOMA_HEADER"
    ],

    boldUnderlineKeys: [
      "REMOVE_MEIBOMIAN_GLAND_ADENOMA",
      "LUBRICATE_MEIBOMIAN_GLAND_ADENOMA"
    ],

    linkKeys: [
      "MEIBOMIAN_GLAND_ADENOMA_ARTICLE"
    ],
    };
    }

  // Nuclear Sclerosis
    function generateCanineNuclearSclerosisTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["NUCLEAR_SCLEROSIS"],
    text: [
    `Nuclear sclerosis: The haziness you’re seeing in your dog’s eyes is due to nuclear sclerosis (also called lenticular sclerosis). This is a normal aging process where the lens, the part of the eye that lets us change our focus from nearby objects to far objects, becomes thicker. This does not impede vision in any way. Your dog can still drive and read just as well as before. You can learn more from the Lenticular Sclerosis in Dogs article on VCA Animal Hospitals.`
    ].join('\n'),

    boldKeys: [
      "NUCLEAR_SCLEROSIS_HEADER"
    ],

    linkKeys: [
      "LENTICULAR_SCLEROSIS_IN_DOGS_ARTICLE"
    ],
    };
    }

/* ------------------ CANINE CARDIOLOGY ------------------ */
  // 2nd Degree AV Block Mobitz Type II | Asymptomatic
    function generateCanine2ndDegreeAVBlockTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["SECOND_DEGREE_AV"],
    text: [
    `2nd Degree Atrioventricular Block: An EKG was performed which shows that your dog has an atrioventricular block. This means that occasionally your dog’s heart will skip a beat. Mobitz type II indicates that this occurs consistently in your dog. Symptoms of this disorder include generalized lethargy, decreased energy when exercising, coughing, or collapse. At this time your dog doesn’t show signs of heart disease. An EKG should be performed every 6 months to ensure there are no changes to your dog’s heart rhythm. In the meantime, continue to monitor for the aforementioned symptoms. Most importantly, count how fast your dog breathes while sleeping. If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease. Contact the clinic immediately.`
    ].join('\n'),

    boldKeys: [
      "SECOND_DEGREE_ATRIOVENTRICULAR_BLOCK_HEADER"
    ],

    boldUnderlineKeys: [
      "EKG_RECOMMENDATION",
      "HEART_MURMUR_WARNING_SIGNS",
    ],
    };
    }

  // Heart Murmur | 0th Discovered, No Tests
    function generateCanineHeartMurmur0Template(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Heart murmur: A heart murmur was heard in your dog today. Heart murmurs are sounds produced whenever blood moves in a direction or location it isn’t meant to. Common causes include heartworms, heart disease, or fetal abnormalities. Grading is based on how loud the sound is. A higher grade (5 & 6) does not always indicate worse disease & a lower grade (1 & 2) does not always indicate a better disease. Diagnosis involves X-rays to see the shape & size of the heart can be performed in clinic and an echocardiogram to look at the inner workings of the heart and find a cause of disease can be scheduled as well. `,
    `In the meantime, monitor your dog for symptoms such as coughing, increased exhaustion when exercising, & low energy. Most importantly, count how fast your dog breathes while sleeping. If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease. Contact the clinic immediately. You can learn more about heart murmurs from the Heart Murmurs in Dogs and Cats article on Veterinary Partner.`
    ].join('\n');


    return {
    sex,
    plurality,
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
    };

  // Heart Murmur | 1st, Normal Radiographs
    function generateCanineHeartMurmur1RadiographsNormalTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Heart murmur: A heart murmur was heard in your dog today. Heart murmurs are sounds produced whenever blood moves in a direction or location it isn’t meant to. Common causes include heartworms, heart disease, or fetal abnormalities. Grading is based on how loud the sound is. A higher grade (5 & 6) does not always indicate worse disease & a lower grade (1 & 2) does not always indicate a better disease.`,
    `X-rays to see the shape & size of the heart were performed and your dog’s heart doesn’t appear to be concerningly enlarged. At this time treatment with medication is not warranted, but an echocardiogram to look at the inner workings of the heart and diagnose the cause of the disease will need to be scheduled.`,
    `In the meantime, monitor your dog for symptoms such as coughing, increased exhaustion when exercising, & low energy. Most importantly, count how fast your dog breathes while sleeping. If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease. Contact the clinic immediately. You can learn more about heart murmurs from the Heart Murmurs in Dogs and Cats article on Veterinary Partner.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Heart Murmur | 1st, Cardiomegaly, Start Pimobendan
    function generateCanineHeartMurmur1CardiomegalyTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Heart murmur: A heart murmur was heard in your dog today. Heart murmurs are sounds produced whenever blood moves in a direction or location it isn’t meant to. Common causes include heartworms, heart disease, or fetal abnormalities. Grading is based on how loud the sound is. A higher grade (5 & 6) does not always indicate worse disease & a lower grade (1 & 2) does not always indicate a better disease.`,
    `Diagnosis includes the x-rays that we performed to see the shape & size of the heart. These x-rays show that the heart is enlarged. It is compressing the lungs and trachea to an extent, so your dog will be started on medication to help improve heart function and slow the progression of disease. An echocardiogram to look at the inner workings of the heart and find a cause of disease will need to be scheduled.`,
    `In the meantime, give the medicine prescribed below as treatment to manage the disease. Monitor your dog for symptoms such as coughing, increased exhaustion when exercising, & low energy. Most importantly, count how fast your dog breathes while sleeping. If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease. Contact the clinic immediately. You can learn more about heart murmurs from the Heart Murmurs in Dogs and Cats article on Veterinary Partner.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Heart Murmur | 3rd, Known
    function generateCanineHeartMurmur3KnownTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Heart murmur: Your dog is known to have a heart murmur. Heart murmurs are sounds produced whenever blood moves in a direction or location it isn’t meant to. Monitor your dog for symptoms such as coughing, increased exhaustion when exercising, & low energy. Most importantly, count how fast your dog breathes while sleeping. If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease. Contact the clinic immediately. You can learn more about heart murmurs from the Heart Murmurs in Dogs and Cats article on Veterinary Partner.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Heartworms | Adulticidal Treatment
    function generateCanineHeartwormsAdulticidalTreatmentTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["HEARTWORMS"],
    text: [
    `Heartworms: Unfortunately your dog has tested positive for heartworms. Treatment is necessary for the next several months & is done in stages.`,
    `1st Month: Doxycycline is given for 4 weeks. This will help remove bacteria that are beneficial to heartworm growth & lowers the risk of side effects caused by dying heartworms. Heartworm prevention is also given to kill immature heartworms & prevent re-infection.`,
    `2nd Month: Doxycycline is completed & heartworm prevention is given for another month to prevent re-infection. Waiting a month allows time for the heartworms to become weaker now that their beneficial bacteria are gone, thus making treatment more effective.`,
    `3rd Month: Medication to kill adult heartworms is injected into the back. This injection is often painful & 30% of dogs are sore or form abscesses afterwards. Abscesses resolve with warm compresses within 1 - 4 weeks, & pain medicine can be sent home as needed. Prednisone (a steroid) is given for a month to control inflammation as the older & weaker heartworms die. The medicine should be tapered as described below. Heartworm prevention is continued.`,
    `4th Month: Only younger, stronger heartworms are left. Your dog will be given another injection of heartworm treatment. In 24 hours, one last injection of heartworm adulticidal is given. Another month of steroids are sent home to further reduce the risk of inflammation. The medicine should be tapered as described below. Heartworm prevention is continued.`,
    `It is vital you cage rest your dog throughout the ENTIRE treatment course & for 2 weeks after the last injection. Failure to do so can cause clots to form in the heart & blood vessels which can be fatal. Watch for coughing, gagging, vomiting, diarrhea, or loss of appetite and contact the clinic if noticed. Excessive sluggishness, respiratory distress, and coughing up blood are signs of an emergency. A repeat heartworm test is advised 9 months after the last injection. You can learn more about heartworm treatment & prevention from the Canine Heartworm Guidelines by the American Heartworm Society.`
    ].join('\n'),

    boldKeys: [
      "FIRST_MONTH_HEADER",
      "SECOND_MONTH_HEADER",
      "THIRD_MONTH_HEADER",
      "FOURTH_MONTH_HEADER",
      "HEARTWORM_DISEASE_HEADER"
    ],

    boldUnderlineKeys: [
      "DOXYCYCLINE_FOUR_WEEKS",
      "HEARTWORMS_EMERGENCY",
      "HEARTWORMS_PREVENTION_GIVEN",
      "MELARSOMINE_INJECTION_SCHEDULE",
      "PREDNISONE_SCHEDULE",
      "PREDNISONE_TAPERING",
      "CONTINUE_HEARTWORM_PREVENTION"
    ],

    redKeys: [
      "HEARTWORMS_CAGE_REST"
    ],

    linkKeys: [
      "CANINE_HEARTWORM_GUIDELINES_BY_THE_AMERICAN_HEARTWORM_SOCIETY_ARTICLE"
    ],
    };
    }

  // Heartworms | Slow Kill Treatment (Healthy Dog)
    function generateCanineHeartwormsSlowKillHealthyDogTreatmentTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["HEARTWORMS"],
    text: [
    `Heartworms: Unfortunately your dog has tested positive for heartworms. Treatment is typically done via the use of injections as this is the most effective and quickest way to treat the disease. However, you’ve chosen to use the “slow kill” method instead. This method is typically reserved for older, delicate patients who may not otherwise survive treatment. It involves using an antibiotic called doxycycline for a month followed by monthly heartworm prevention.`,
    `Note that it can take years for heartworms to be completely eradicated via the slow kill method. There is still a chance of death during this time. Additionally, your dog’s internal organs are still at risk of getting damaged from the heartworms, and exercise restriction is strongly advised until a negative test is obtained. You can learn more about the slow kill method from the American Heartworm Society statement. More information about heartworm treatment & prevention is available from the Canine Heartworm Guidelines by the American Heartworm Society.`
    ].join('\n'),

    boldKeys: [
      "HEARTWORM_DISEASE_HEADER"
    ],

    boldUnderlineKeys: [
      "HEARTWORMS_SLOW_KILL_WARNING"
    ],

    linkKeys: [
      "CANINE_HEARTWORM_GUIDELINES_BY_THE_AMERICAN_HEARTWORM_SOCIETY_ARTICLE",
      "THE_AMERICAN_HEARTWORM_SOCIETY_STATEMENT_ARTICLE"
    ],
    };
    }

  // Heartworm Test Repeat
    function generateCanineHeartwormTestRepeatTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    rank: 998,
    cleanupKeys: ["HEARTWORMS_PREVENTION_HEADER"],
    text: [
    `Heartworms: Heartworms are spread by mosquitoes which don’t die in the Texas "winter", so our pets are at risk of infection year round. It takes 6 months for heartworms to grow from the “baby” form to the adult form, and our tests only detect the adult form. For that reason it’s important your ${g.dog} ${g.comes} back in 6 months for a repeat of the heartworm test.`,
    `Furthermore, heartworms can be fatal & there is a risk of death even with proper treatment (which requires staying in a cage for 6 months and three painful injections). Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your dog on monthly prescription preventatives such as Heartgard, Simparica Trio, Revolution, etc.`,
    `After the repeat heartworm test in 6 months, the Proheart injection can be given to protect against heartworms for an entire year. You can learn about the different types of heartworm prevention from the Canine Heartworms and Preventing Disease article, the Preventing Heartworm Infection in Dogs article, and the Heartworm Prevention Comparison Chart for Dogs and Cats article on Veterinary Partner.`
    ].join('\n'),
    
    boldKeys: [
      "HEARTWORM_DISEASE_HEADER"
    ],

    boldUnderlineKeys: [
      "HEARTWORM_TEST_REPEAT",
    ],

    linkKeys: [
      "CANINE_HEARTWORMS_AND_PREVENTING_DISEASE_ARTICLE",
      "HEARTWORM_PREVENTION_COMPARISON_CHART_FOR_DOGS_AND_CATS_ARTICLE",
      "PREVENTING_HEARTWORM_INFECTION_IN_DOGS_ARTICLE"
    ],
    };
    }

  // Hypertension | 1st, Diagnosed
    function generateCanineHypertension1DiagnosedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["HYPERTENSION"],
    text: [
    `Hypertension: Your dog has elevated blood pressure. Common causes include a variety of diseases such as chronic kidney failure, kidney damage, diabetes, or hyperadrenocorticism (Cushing’s disease) to name a few. Regardless of the cause, symptoms include progressive blindness (bumping into walls/furniture, barking when not held, etc.), damage to the kidneys (pain/yelping when urinating, more frequent urination, blood in the urine, etc.), and the formation of dangerous blood clots (may not show symptoms but can life threatening).`,
      `Diagnosis was through the blood pressure test that has already been done. Treatment involves a medication called amlodipine. The medicine dilates the blood vessels so that pressure drops. Give consistently as prescribed below and bring your dog back in 1 week for a recheck of blood pressure. You can learn more about this disease from the High Blood Pressure in Our Pets article on Veterinary Partner. This is a lifelong disease which can be managed but not cured. Your dog will PERMANENTLY require medication.`
    ].join('\n'),

    boldKeys: [
      "HYPERTENSION_HEADER"
    ],

    boldUnderlineKeys: [
      "HYPERTENSION_DIAGNOSIS",
      "HYPERTENSION_MEDICINE",
    ],

    greenKeys: [
      "COMMON_CAUSES",
      "DIAGNOSIS",
      "TREATMENT",
      "SYMPTOMS"
    ],

    linkKeys: [
      "HIGH_BLOOD_PRESSURE_IN_OUR_PETS_ARTICLE"
    ],
    };
    }


  // Left Sided Congestive Heart Failure
    function generateCanineLeftSidedCongestiveHeartFailureTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["LEFT_SIDED_CONGESTIVE_HEART_FAILURE"],
    text: [
    `Left sided congestive heart failure: Your dog has been diagnosed with congestive heart failure. Common causes include mitral valve disease, dilated cardiomyopathies, and most congenital heart defects to name a few. Regardless of the cause, fluid builds up in the lungs and the space surrounding them since the heart isn't sending the blood to the rest of the body as it should.`,
      `Diagnosis was achieved through a combination of x-rays, echocardiogram, and physical exam findings. Treatment involves using medicine to make it easier for the heart to pump, using medicine to remove fluid around the heart and lungs, & possibly surgery to fix blood going the wrong direction depending on the cause.`,
      `For now, monitor your dog for symptoms such as coughing, increased exhaustion when exercising, & low energy. Most importantly, count how fast your dog breathes while sleeping. If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease. Contact the clinic immediately. You can learn more about congestive heart failure from the Congestive Heart Failure in Dogs & Cats article on Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "LEFT_SIDED_CONGESTIVE_HEART_FAILURE_HEADER"
    ],

    boldUnderlineKeys: [
      "HEART_MURMUR_WARNING_SIGNS",
    ],

    greenKeys:
    [
      "COMMON_CAUSES",
      "DIAGNOSIS",
      "SYMPTOMS",
      "TREATMENT",
    ],

    linkKeys: [
      "CONGESTIVE_HEART_FAILURE_IN_DOGS_CATS_ARTICLE_ARTICLE"
    ],
    };
    }

  // Myxomatous Mitral Valve Disease | Diagnosed
    function generateCanineMyxomatousMitralValveDiseaseTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["MYXOMATOUS_MITRAL_VALVE_DISEASE"],
    text: [
    `Myxomatous mitral valve disease: The heart murmur present in your dog is caused by a normal aging process called myxomatous mitral valve disease. One of the heart valves that stops blood from going backwards is no longer doing its job with enthusiasm. This is a common finding in older dogs & usually doesn’t cause illness.`,
    `Monitor your dog for signs of coughing, increased exhaustion when exercising, & low energy. Most importantly, count how fast your dog breathes while sleeping. If you notice a respiratory rate above 35 breaths per minute while sleeping or any of the other signs, these may indicate worsening heart disease. Contact the clinic immediately. Chest x-rays and an echocardiogram are recommended every 6 - 12 months. You can learn more from the Mitral Valve Disease in Dogs article on Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "MYXOMATOUS_MITRAL_VALVE_DISEASE_HEADER"
    ],

    boldUnderlineKeys: [
      "HEART_MURMUR_WARNING_SIGNS",
    ],

    linkKeys: [
      "MITRAL_VALVE_DISEASE_IN_DOGS_ARTICLE"
    ],
    };
    }

/* ------------------ CANINE RESPIRATORY ------------------ */
  // Bordetellosis | 0th, Presumed
    function generateCanineBordetellosis0PresumedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["BORDETELLOSIS_PRESUMED"],
    text: [
    `Bordetellosis: Based on the symptoms seen in clinic, your dog most likely has bordetellosis, more commonly called kennel cough. Bordetella is a bacteria that is spread through the air whenever a dog enters an area that an infected dog has been to, including dog parks, grooming facilities, & veterinary clinics. Approximately 90% of kennel cough cases resolve on their own, so we will be using cough tablets to help your dog get through this infection.`,
    `Bring your dog back in 1 week for a recheck appointment if no improvement is seen (return immediately if worsening). Symptoms to watch for include lethargy, worsening cough, and decreased appetite. In the meantime keep your dog away from other dogs (such as grooming facilities, boarding locations, and dog parks) until a few days after the cough has stopped. You can learn more about bordetella from the Kennel Cough in Dogs article by Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "BORDETELLOSIS_HEADER"
    ],

    boldUnderlineKeys: [
      "RECHECK_ADVISE_1_WEEK"
    ],

    italicKeys: [
      "BORDETELLA_DISEASE"
    ],

    linkKeys: [
      "BORDETELLA_LINK"
    ],
    };
    }

  // Brachycephalic Obstructive Airway Syndrome
    function generateCanineBrachycephalicObstructiveAirwaySyndromeTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["BRACHYCEPHALIC_OBSTRUCTIVE_AIRWAY_SYNDROME"],
    text: [
    `Brachycephalic obstructive airway syndrome: Your ${g.dogss} breed is known to have a disorder called brachycephalic obstructive airway syndrome. There are multiple parts to this disease, but the two most common problems are elongated soft palate & stenotic nares (closed nostrils). The soft palate is equivalent to the human uvula. It extends into the windpipe & causes difficulty breathing & overheating. This also causes water & food to enter the windpipe which may lead to pneumonia. Treatment is using surgery to remove the elongated soft palate.`,
    `Stenotic nares means that the nose doesn’t open fully when breathing. This causes difficulty breathing, overheating, & excessive panting. Treatment is surgery to open up the nares more. These surgeries are often done together & can increase the quality and length of your ${g.dogss} ${g.life}.`,
    `Overheating: If you are ever concerned your ${g.dog} ${g.is} getting too hot, you can have ${g.him} lay on a tile or wood floor and place a towel soaked with cool water on top. DO NOT pour cold water or ice on ${g.him} or the towel as this can cause the blood vessels to constrict, thereby putting your ${g.dog} in a state of shock. Direct a fan towards the towel to further increase cooling. During hot days, limit outside exposure to no more than 5 minutes. Freeze a block of ice and place it in the water bowl. You can use a plastic container or clean out food tubs (such as ice cream gallon tubs) to make the ice block. Place a small amount of water in the tube so that the water slowly melts throughout the day and provides cool water at all times.`
    ].join('\n'),
    
    boldKeys: [
      "BRACHYCEPHALIC_OBSTRUCTIVE_AIRWAY_SYNDROME_HEADER",
      "OVERHEATING_HEADER"
    ],

    boldUnderlineKeys: [
      "BOAS_GENETICS",
      "ELONGATED_SOFT_PALATE_TX",
      "STENOTIC_NARES_TX",
    ],

    redKeys: [
      "BOAS_COLD_WATER_SHOCK"
    ],
    };
    }
  
  // Chronic Bronchitis | 0th, Presumed
    function generateCanineChronicBronchitis0PresumedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["CHRONIC_BRONCHITIS_PRESUMED"],
    text: [
    `Chronic bronchitis: Your dog appears to have chronic bronchitis, meaning some of the main airways (the bronchi) have inflammation in them. Common causes include airway irritants (air pollution, smoke in the house from cooking or otherwise, burning incense, etc.) though the normal aging process can also contribute. The resulting inflammation causes overproduction of mucus which blocks the airway, causing symptoms such as coughing as the lungs try to clear it away. Once mucus production starts, it is difficult to completely eradicate.`,
    `Diagnosis requires x-rays to make sure the cough isn’t being caused by heart problems or other lung diseases such as tracheal collapse or pneumonia. Treatment involves managing the disease through weight loss and using medicine to either open up the airways, reduce inflammation, or treat present infection. Give the medicine prescribed below as directed. You can learn more from the Chronic Bronchitis in Dogs article on Veterinary Partner.`,
    `Bring your dog back in 1 week for a recheck appointment if no improvement is seen (return immediately if worsening). Otherwise contact the clinic if the medication does not prove effective within 2 weeks.`
    ].join('\n'),

    boldKeys: [
      "CHRONIC_BRONCHITIS_HEADER"
    ],

    boldUnderlineKeys: [
      "RECHECK_ADVISE_1_WEEK"
    ],

    greenKeys: [
      "COMMON_CAUSES",
      "DIAGNOSIS",
      "SYMPTOMS",
      "TREATMENT"
    ],

    linkKeys: [
      "CHRONIC_BRONCHITIS_IN_DOGS_ARTICLE_ARTICLE"
    ],
    };
    }
  
  // Collapsing Trachea | 1st, Theophylline
    function generateCanineCollapsingTrachea1TheophyllineTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["COLLAPSING_TRACHEA"],
    text: [
    `Collapsing trachea: Your dog appears to have a collapsing trachea. This occurs when the cartilage rings in the trachea weaken and cause the trachea (the windpipe) to flatten whenever your dog breathes in. The result is that your dog has difficulty breathing and may even start coughing, both of which can cause your dog distress and discomfort. Common causes of tracheal collapse include obesity, respiratory irritants such as smoke, incense, dust, etc.), and heart enlargement.`,
      `Treatment involves medication to make it easier for your dog to open up the airways and make it easier for your dog to breathe. This will decrease how deeply your dog needs to breathe with each breath which makes the trachea collapse less. Surgery can be pursued to keep the airways open longer and referral to a surgeon for that can be discussed if necessary. You can learn more from the Tracheal Collapse in Dogs article on Veterinary Partner. `
    ].join('\n'),

    boldKeys: [
      "COLLAPSING_TRACHEA_HEADER"
    ],

    greenKeys: [
      "COMMON_CAUSES",
      "TREATMENT"
    ],

    linkKeys: [
      "TRACHEAL_COLLAPSE_IN_DOGS_ARTICLE_ARTICLE"
    ],

    };
    }

  // Collapsing Trachea | 3rd, Known, No Meds
    function generateCanineCollapsingTrachea3NoMedsTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["COLLAPSING_TRACHEA"],
    text: [
    `Collapsing trachea: Your dog is known to have a collapsing trachea. This occurs when the cartilage rings in the trachea weaken and cause the trachea (the windpipe) to flatten whenever your dog breathes in. The result is that your dog has difficulty breathing and may even start coughing, both of which can cause your dog distress and discomfort. Common causes of tracheal collapse include obesity, respiratory irritants such as smoke, incense, dust, etc.), and heart enlargement.`,
    `Treatment involves medication to make it easier for your dog to open up the airways and make it easier for your dog to breathe. This will decrease how deeply your dog needs to breathe with each breath which makes the trachea collapse less. At this time no medication is being used as your dog appears to be well controlled at home without it. Surgery can be pursued to keep the airways open longer and referral to a surgeon for that can be discussed if necessary. You can learn more from the Tracheal Collapse in Dogs article on Veterinary Partner. `
    ].join('\n'),

    boldKeys: [
      "COLLAPSING_TRACHEA_HEADER"
    ],

    boldUnderlineKeys: [
      "COLLAPSING_TRACHEA_NO_TREATMENT"
    ],

    greenKeys: [
      "COMMON_CAUSES",
      "TREATMENT"
    ],

    linkKeys: [
      "TRACHEAL_COLLAPSE_IN_DOGS_ARTICLE_ARTICLE"
    ],
    };
    }

  // Laryngeal Paralysis | 1st, Diagnosed
    function generateCanineLaryngealParalysis1DiagnosedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["LARYNGEAL_PARALYSIS"],
    text: [
    `Laryngeal paralysis: The roaring sound present whenever your dog breathes is caused by laryngeal paralysis. When we breathe in, the larynx (commonly called the voice box) opens up to allow air to enter the trachea & then the lungs. Dogs with laryngeal paralysis have part of the larynx paralyzed. This means that it doesn’t fully open up & breathing becomes more difficult.`,
    `Symptoms include difficulty breathing, panting excessively, or a voice change. You can help improve your dog’s ability to breathe by switching from a collar to a harness, reducing activity, & using anti-anxiety medication if anxiety is present. Alternatively, surgery can be performed to keep the larynx open permanently. You can learn more about laryngeal paralysis from the Laryngeal Paralysis in Dogs article from Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "LARYNGEAL_PARALYSIS_HEADER"
    ],

    boldUnderlineKeys: [
      "You can help improve your dog’s ability to breathe by switching from a collar to a harness, reducing activity, & using anti-anxiety medication if anxiety is present."
    ],

    greenKeys: [
      "SYMPTOMS"
    ],

    linkKeys: [
      "LARYNGEAL_PARALYSIS_IN_DOGS_ARTICLE_ARTICLE"
    ],
    };
    }

  // Reverse Sneezing
    function generateCanineReverseSneezingTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["REVERSE_SNEEZING"],
    text: [
    `Reverse sneezing: The sound your dog produces is known as a reverse sneeze. This is caused by irritation in the part of the mouth that connects to the nose. You may see this if something irritates your dog's throat (such as perfume, smoke, food, etc.), after excitement, pulling on the leash, or allergies. Treatment includes massaging the throat when it occurs or and giving Benadryl 25mg (give 1 tablet per 25 lbs every 12 hours) or Zyrtec 10mg (give up to 1 tablet per 10 lbs every 12 hours) for allergies. Side effects (drowsiness, thirst) are more common with Benadryl than Zyrtec.`
    ].join('\n'),

    boldKeys: [
      "REVERSE_SNEEZING_HEADER"
    ],

    boldUnderlineKeys: [
      "REVERSE_SNEEZING_TREATMENT"
    ],

    greenKeys: [
      "TREATMENT"
    ],
    };
    }

/* ------------------ CANINE ENDOCRINE ------------------ */
  // Diabetes Mellitus | 1st, Diagnosed
    function generateCanineDiabetesMellitus1DiagnosedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["DIABETES_MELLITUS"],
    text: [
    `Diabetes mellitus: Based on physical findings and bloodwork results, your dog suffers from diabetes. In dogs this often occurs because their body isn't producing enough insulin. As such, glucose (sugar) isn't being taken out of ${g.his} blood and into ${g.his} cells where they're needed. Symptoms of diabetes include increased thirst, urination, appetite, weight loss, and seizures. Treatment requires giving insulin injections every 12 hours after food, and you are advised to switch to a high fiber/low fat diet. Give insulin AFTER your pet has eaten. Do not mix up 100 unit insulin syringes and 40 unit insulin syringes.`,
    `If you are concerned your dog got too much insulin or got a second helping, you can give 1 tablespoon of light Karo syrup or honey per 5 lbs. If signs persist, seek immediate medical treatment. Your dog will need to come back in two weeks for a glucose curve. You can learn more about diabetes in pets from the Diabetes Mellitus Introduction article on Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "DIABETES_HEADER"
    ],

    boldUnderlineKeys: [
      "DIABETES_SYMPTOMS",
      "GLUCOSE_CURVE_ADVISED"
    ],

    greenKeys: [
      "SYMPTOMS",
      "TREATMENT"
    ],

    redKeys: [
      "INSULIN_ADMINISTRATION_WARNING",
      "DIABETES_EMERGENCY",
    ],

    linkKeys: [
      "DIABETES_MELLITUS_INTRODUCTION_ARTICLE",
      "LIGHT_KARO_SYRUP_ARTICLE"
    ],
    };
    }

  // Diabetes Mellitus | 3rd, Controlled
    function generateCanineDiabetesMellitus3ControlledTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["DIABETES_MELLITUS"],
    text: [
    `Diabetes mellitus: Your dog is known to have diabetes which appears to be well controlled with your current insulin amount. Continue to monitor for signs of uncontrolled diabetes such as increased thirst, urination, appetite, weight loss, and seizures. Give insulin AFTER your pet has eaten. Do not mix up 100 unit insulin syringes and 40 unit insulin syringes. If you are concerned your dog got too much insulin or got a second helping, you can give 1 tablespoon of light Karo syrup or honey per 5 lbs. If signs persist, seek immediate medical treatment. You can learn more about diabetes in pets from the Diabetes Mellitus Introduction article on Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "DIABETES_HEADER"
    ],

    boldUnderlineKeys: [
      "CONTINUE_MONITOR_DIABETES"
    ],

    redKeys: [
      "DIABETES_EMERGENCY",
      "INSULIN_ADMINISTRATION_WARNING"
    ],

    linkKeys: [
      "DIABETES_MELLITUS_INTRODUCTION_ARTICLE",
      "LIGHT_KARO_SYRUP_ARTICLE"
    ],
    };
    }

  // Hypothyroidism | 1st, Diagnosed
    function generateCanineHypothyroidism1DiagnosedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["Hypothyroidism"],
    text: [
    `Hypothyroidism: Your dog has been diagnosed with hypothyroidism. Hypothyroidism in dogs occurs when the thyroid gland doesn’t produce as much hormone as it should due to the body’s own immune system attacking the thyroid gland. This causes decreased energy, increased weight gain, & dry hair coat. Treatment involves giving thyroid hormone as a pill.`,
    `We will need to recheck thyroid hormone levels in 4 - 6 weeks to ensure we’re not giving too much thyroid & causing hyperthyroidism. Give the medicine 4 - 6 hours BEFORE the appointment. Medication must be given consistently as prescribed. You can learn more about hypothyroidism from the Hypothyroidism in Dogs article by Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "HYPOTHYROIDISM_HEADER"
    ],

    boldUnderlineKeys: [
      "HYPOTHYROIDISM_DIAGNOSED",
      "HYPOTHYROID_RECHECK",
    ],

    linkKeys: [
      "HYPOTHYROIDISM_IN_DOGS_ARTICLE_ARTICLE"
    ],
    };
    }

  // Hypothyroidism | 2nd, Recheck
    function generateCanineHypothyroidism2RecheckTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["HYPOTHYROIDISM"],
    text: [
    `Hypothyroidism: Your dog is known to have hypothyroidism, a disease process where the body’s own immune system attacks the thyroid gland, thereby decreasing the amount of thyroid hormone produced. You are also giving a thyroid hormone supplement pill to counteract that reduction. Samples were drawn from your dog today to check the current thyroid hormone levels. You will be called with results in 3 - 4 business days. You can learn more about hypothyroidism from the Hypothyroidism in Dogs article by Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "HYPOTHYROIDISM_HEADER"
    ],

    boldUnderlineKeys: [
      "YOU_WILL_BE_CALLED_WITH_RESULTS"
    ],

    linkKeys: [
      "HYPOTHYROIDISM_IN_DOGS_ARTICLE_ARTICLE"
    ],
    };
    }

  // Hypothyroidism | 3rd, Controlled
    function generateCanineHypothyroidism3ControlledTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["HYPOTHYROIDISM"],
    text: [
    `Hypothyroidism: Your dog is known to have hypothyroidism. Hypothyroidism in dogs occurs when the thyroid gland doesn’t produce as much hormone as it should due to the body’s own immune system attacking the thyroid gland, leading to decreased energy, increased weight gain, & dry hair coat. Bloodwork was performed which shows that the thyroid hormone levels are within normal limits. Continue to give the medication as you have been. Bloodwork is recommended every 6 months. You can learn more about hypothyroidism from the Hypothyroidism in Dogs article by Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "HYPOTHYROIDISM_HEADER"
    ],

    boldUnderlineKeys: [
      "CONTINUE_MEDICATION_LABWORK_6_MONTHS",
    ],

    linkKeys: [
      "HYPOTHYROIDISM_IN_DOGS_ARTICLE_ARTICLE"
    ],
    };
    }

  // Hyperadrenocorticism | 1st, ACTH Stim Test
    function generateCanineHyperadrenocorticism1ACTHStimTestTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["HYPERADRENOCORTICISM_PRESUMED"],
    text: [
    `Hyperadrenocorticism: Your dog shows signs of hyperadrenocorticism (also known as HOC, Cushing’s disease, or Cushing’s syndrome). The most common causes of this disease is due to a tumour in the brain or adrenal glands causing an overproduction of cortisol, the stress hormone. Symptoms of the disease include hair loss, increased thirst/urination, voracious appetite, enlarged belly, lethargy, recurrent infections, and excessive panting.`,
    ` In order to diagnose the disease, the ACTH stim test was performed and blood has been drawn. You will be called with results in 3 - 4 business days. This test only tells us if your dog has hyperadrenocorticism. It does not tell us whether the tumour is in the brain or adrenal glands. Depending on the location, treatment with medicine (such as trilostane) or surgery (adrenalectomy) to remove part or all of the adrenal gland can be discussed.`,
    `If trilostane is given, a recheck ACTH stim test should be performed two weeks after starting meds to make sure your dog is receiving the right amount. You can learn more about hyperadrenocorticism from the Cushing's Syndrome (Hyperadrenocorticism) article on Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "HYPERADRENOCORTICISM_HEADER"
    ],

    boldUnderlineKeys: [
      
      "RECHECK_ACTH_STIM_TEST",
      "YOU_WILL_BE_CALLED_WITH_RESULTS"
    ],

    greenKeys: [
      "COMMON_CAUSES",
      "DIAGNOSE",
      "SYMPTOMS",
      "TREATMENT",
    ],

    linkKeys: [
      "CUSHINGS_SYNDROME_HYPERADRENOCORTICISM_ARTICLE"
    ],
    };
    }

  // Hyperadrenocorticism | 2nd, Diagnosed
    function generateCanineHyperadrenocorticism2DiagnosedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["HYPERADRENOCORTICISM"],
    text: [
    `Hyperadrenocorticism: Your dog has been diagnosed with hyperadrenocorticism (also known as HOC, Cushing’s disease, or Cushing’s syndrome). We will be treating using trilostane to prevent the overproduction of the steroid hormone known as cortisol. However, it’s important we don’t cause an underproduction of cortisol. As such we’ll be starting your dog on a low dose and increasing to a higher dose as needed.`,
    `We will need to perform an ACTH stim test in 2 weeks to ensure your dog’s cortisol levels are normal. Make sure to give the trilostane 4 - 6 hours before bringing your dog into the clinic. If your dog is feeling lethargic, vomiting, or has a decreased appetite, contact the clinic immediately and we can discuss having the test done sooner. You can learn more about hyperadrenocorticism from the Cushing's Syndrome (Hyperadrenocorticism) article on Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "HYPERADRENOCORTICISM_HEADER"
    ],

    boldUnderlineKeys: [
      "TRILOSTANE_HYPOADRENOCORTICISM",
      "TRILOSTANE_WARNING",
    ],

    linkKeys: [
      "CUSHINGS_SYNDROME_HYPERADRENOCORTICISM_ARTICLE"
    ],
    };
    }

  // Hyperadrenocorticism | 3rd, Controlled
    function generateCanineHyperadrenocorticism3ControlledTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["HYPERADRENOCORTICISM"],
    text: [
    `Hyperadrenocorticism: Your dog is known to have hyperadrenocorticism (also known as HOC, Cushing’s disease, or Cushing’s syndrome) and is currently on trilostane to control it. The ACTH stim is recommended every 3 - 4 months to ensure the disease is still well monitored. Adjustments to the dose can be made as needed. Continue to monitor for side effects including vomiting, diarrhea, listlessness, or decreased water intake. You can learn more about hyperadrenocorticism from the Cushing's Syndrome (Hyperadrenocorticism) article on Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "HYPERADRENOCORTICISM_HEADER"
    ],

    boldUnderlineKeys: [
      "HYPERADRENOCORTICISM_SYMPTOMS"
    ],

    linkKeys: [
      "CUSHINGS_SYNDROME_HYPERADRENOCORTICISM_ARTICLE"
    ],
    };
    }

  // Hyperadrenocorticism | 4th, Uncontrolled
    function generateCanineHyperadrenocorticism4UncontrolledTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["HYPERADRENOCORTICISM"],
    text: [
    `Hyperadrenocorticism: Your dog is known to have hyperadrenocorticism (also known as HOC, Cushing’s disease, or Cushing’s syndrome) and is currently on trilostane to control it. The ACTH stim test was performed today which shows that the disease is not fully controlled with the current medication. We will need to increase by 25%. Continue to monitor for side effects including vomiting, diarrhea, listlessness, or decreased water intake. A repeat test is necessary 7 days after starting the new medication. You can learn more about hyperadrenocorticism from the Cushing's Syndrome (Hyperadrenocorticism) article on Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "HYPERADRENOCORTICISM_HEADER"
    ],

    boldUnderlineKeys: [
      "HYPERADRENOCORTICISM_REPEAT_TEST"
    ],

    linkKeys: [
      "CUSHINGS_SYNDROME_HYPERADRENOCORTICISM_ARTICLE"
    ],
    };
    }

  // Hyperadrenocorticism | 5th, Regular Checkup
    function generateCanineHyperadrenocorticism5CheckupTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["HYPERADRENOCORTICISM"],
    text: [
    `Hyperadrenocorticism: Your dog is known to have hyperadrenocorticism (also known as HOC, Cushing’s disease, or Cushing’s syndrome) and is currently on trilostane to control it. The ACTH stim test was performed today to make sure your dog’s disease is well controlled. If abnormalities occur, a change in medication will be necessary. You will be called with results in 3 - 4 business days. Until then, continue your current medication as previously prescribed. If the disease is being well controlled with medication, a repeat ACTH stim test and bloodwork is advised every 3 - 4 months. You can learn more about hyperadrenocorticism from the Cushing's Syndrome (Hyperadrenocorticism) article on Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "HYPERADRENOCORTICISM_HEADER"
    ],

    boldUnderlineKeys: [
      "CONTINUE_MEDICATION_AS_PRESCRIBED",
      "YOU_WILL_BE_CALLED_WITH_RESULTS"
    ],

    linkKeys: [
      "CUSHINGS_SYNDROME_HYPERADRENOCORTICISM_ARTICLE"
    ],
    };
    }

  // Pancreatitis | 1st, Diagnosed
    function generateCaninePancreatitis1DiagnosedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["PANCREATITIS"],
    text: [
    `Pancreatitis: The canine pancreatic lipase immunoreactivity (cPLI) test was performed today which shows that your dog has pancreatitis. This is an inflammation of the pancreas that is caused by a variety of different issues. Symptoms include decreased appetite, vomiting, diarrhea, painful abdomen, and excessive panting/sweating from the paw pads. The best way to prevent pancreatic flare ups is by feeding a low fat diet and cutting out treats that are high in fat. You can learn more about pancreatitis from the Pancreatitis in Dogs article on Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "PANCREATITIS_HEADER"
    ],

    boldUnderlineKeys: [
      "PANCREATITIS_PREVENTION"
    ],

    greenKeys: [
      "SYMPTOMS"
    ],

    linkKeys: [
      "PANCREATITIS_IN_DOGS_ARTICLE"
    ],
    };
    }

  // Pancreatitis | 2nd, Flare Up
    function generatecCaninePancreatitis2FlareUpTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["PANCREATITIS"],
    text: [
    `Pancreatic flare up: Your dog is known to have pancreatitis or has previously had it & is currently suffering from a flare up. These flare ups can occur due to stress, hormone imbalance, trauma, & high fat meals. In order to reduce pain & discomfort related to pancreatitis, pain medicine (gabapentin), anti-vomiting/anti-nausea medicine (Cerenia), & fluids are recommended. Sucralfate, a medicine that covers ulcers & prevents further intestinal pain, has also been prescribed.`,
    `Going forward, you can help reduce the risk of pancreatitis by removing any source of stress for your dog & feeding low fat diets such as Hill’s Science Diet i/d low fat (i/d low fat dry food, i/d low fat stew, or i/d low fat patée), Purina EN low fat (EN low fat dry food or EN low fat wet food), or Royal Canine Gastrointestinal Low Fat (RC small dog dry, RC large dog dry, or RC wet food). You can learn more about pancreatitis from the Pancreatitis in Dogs article on Veterinary Partner.`
    ].join('\n'),

    boldKeys: [
      "PANCREATIC_FLARE_UP_HEADER"
    ],

    linkKeys: [
      "EN_LOW_FAT_DRY_FOOD_ARTICLE",
      "EN_LOW_FAT_WET_FOOD_ARTICLE",
      "I_D_LOW_FAT_DRY_FOOD_ARTICLE",
      "I_D_LOW_FAT_PAT_E_ARTICLE",
      "I_D_LOW_FAT_STEW_ARTICLE",
      "PANCREATITIS_IN_DOGS_ARTICLE",
      "RC_LARGE_DOG_DRY_ARTICLE",
      "RC_SMALL_DOG_DRY_ARTICLE",
      "RC_WET_FOOD_ARTICLE"
    ],
    };
    }

/* ------------------ CANINE GASTROINTESTINAL ------------------ */
  // Acute Gastroenteritis | Diarrhea, Home Diet
    function generateCanineAcuteGastroenteritisDiarrheaHomeDietTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["ACUTE_GASTROENTERITIS_DIARRHEA_ONLY"],
    text: [
    `Acute gastroenteritis: At this time no overt causes of diarrhea were identified. Based on your dog’s history & age, the most likely cause of diarrhea is dietary indiscretion (eating something that isn’t healthy for dogs). Ideally your dog would be fed a prescription gastrointestinal diet as a bland, easy to digest aid. At this time you’ve elected to use a homemade bland diet of boiled chicken & rice without salt or other spices. You can also add on psyllium husk (½ gram per lb) once daily for further fiber support. Anti-diarrheal medicine has been sent home for the next two weeks.`
    ].join('\n'),

    boldKeys: [
      "ACUTE_GASTROENTERITIS_HEADER"
    ],
    };
    }

  // Acute Gastroenteritis | Diarrhea, Fecal Test
    function generateCanineAcuteGastroenteritisDiarrheaFecalTestTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["ACUTE_GASTROENTERITIS_DIARRHEA_ONLY"],
    text: [
    `Acute gastroenteritis: At this time no overt causes of diarrhea were identified. A fecal test to check for intestinal parasites is currently running. You will be called with results in 3 - 4 business days. In the meantime your dog’s diarrhea will be treated symptomatically. Feed the prescription gastrointestinal diet as prescribed to speed up the healing process. You can also add on psyllium husk (½ gram per lb) once daily for further fiber support. Anti-diarrheal medicine has been sent home for the next two weeks. `
    ].join('\n'),

    boldKeys: [
      "ACUTE_GASTROENTERITIS_HEADER"
    ],

    boldUnderlineKeys: [
      "YOU_WILL_BE_CALLED_WITH_RESULTS"
    ],
    };
    }

  // Acute Gastroenteritis | Diarrhea, Fecal Test Declined
    function generateCanineAcuteGastroenteritisDiarrheaDeclinedFecalTestTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["ACUTE_GASTROENTERITIS_DIARRHEA_ONLY"],
    text: [
    `Acute gastroenteritis: At this time no overt causes of diarrhea were identified. A fecal test to check for intestinal parasites has been declined so we are treating symptomatically instead. Feed the prescription gastrointestinal diet as prescribed to speed up the healing process. You can also add on psyllium husk (½ gram per lb) once daily for further fiber support. Anti-diarrheal medicine has been sent home for the next two weeks. `
    ].join('\n'),

    boldKeys: [
      "ACUTE_GASTROENTERITIS_HEADER"
    ],
    };
    }

  // AG | Diarrhea/Vomiting, Bldwrk Normal, Fecal Pending
    function generateCanineAGDiarrheaVomitingBloodworkFecalTestTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["ACUTE_GASTROENTERITIS_VOMITING_DIARRHEA"],
    text: [
    `Acute gastroenteritis: At this time no overt causes of vomiting and soft stool were identified. Your dog likely has dietary indiscretion (eating something that isn’t healthy for dogs). Bloodwork to check internal organs came back within normal limits meaning your dog likely hasn’t eaten anything toxic. A fecal test to check for intestinal parasites is still pending. You will be called with results in 3 - 4 business days. In the meantime, medication to help improve the symptoms has been started. Bring your dog back in 3 days for a recheck appointment if no improvement is seen (return immediately if worsening).`
    ].join('\n'),

    boldKeys: [
      "ACUTE_GASTROENTERITIS_HEADER"
    ],

    boldUnderlineKeys: [
      "RECHECK_ADVISE_3_DAYS",
      "YOU_WILL_BE_CALLED_WITH_RESULTS"
    ],
    };
    }

  // Acute Gastroenteritis | Vomiting, Bloodwork Normal
    function generateCanineAcuteGastroenteritisVomitingBloodworkNormalTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["ACUTE_GASTROENTERITIS_VOMITING_ONLY"],
    text: [
    `Acute gastroenteritis: At this time no overt causes of vomiting were identified. Your dog likely has dietary indiscretion (eating something that isn’t healthy for dogs) which led to inflammation of the stomach and intestines (gastroenteritis). Bloodwork was performed which showed normal values meaning whatever it was isn’t toxic enough to damage the liver, kidneys, and other vital internal organs. We will start out by treating your dog’s symptoms. Bring your dog back in 3 days for a recheck appointment if no improvement is seen (return immediately if worsening).`
    ].join('\n'),

    boldKeys: [
      "ACUTE_GASTROENTERITIS_HEADER"
    ],

    boldUnderlineKeys: [
      "RECHECK_ADVISE_3_DAYS"
    ],
    };
    }

  // Acute Gastroenteritis | Vomiting, Bloodwork Declined
    function generateCanineAcuteGastroenteritisVomitingBloodworkDeclinedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["ACUTE_GASTROENTERITIS_VOMITING_ONLY"],
    text: [
    `Acute gastroenteritis: At this time no overt causes of vomiting were identified. Your dog likely has dietary indiscretion (eating something that isn’t healthy for dogs). Bloodwork to check for signs of toxicity or organ dysfunction has been declined, so we will move forward with symptomatic treatment. Medicine to help your dog’s clinical signs improve has been started. Bring your dog back in 3 days for a recheck appointment if no improvement is seen (return immediately if worsening).`
    ].join('\n'),

    boldKeys: [
      "ACUTE_GASTROENTERITIS_HEADER"
    ],

    boldUnderlineKeys: [
      "RECHECK_ADVISE_3_DAYS"
    ],
    };
    }

  // Acute Gastroenteritis | Vomiting, Radiographs
    function generateCanineAcuteGastroenteritisVomitingRadiographsNormalTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["ACUTE_GASTROENTERITIS_VOMITING_ONLY"],
    text: [
    `Acute gastroenteritis: No overt causes of vomiting were identified on physical exam, and x-rays did not find an object stuck in your dog’s stomach or intestines. The most likely cause of vomiting is dietary indiscretion (eating something that isn’t healthy for dogs). Your dog was given an anti-vomiting injection and will need anti-vomiting medicine over the next several days as prescribed below. If you still see vomiting within 24 hours of the injection, your dog needs to go to your nearest veterinary emergency hospital immediately.`
    ].join('\n'),

    boldKeys: [
      "ACUTE_GASTROENTERITIS_HEADER"
    ],

    boldUnderlineKeys: [
      "VOMITING_POST_MAROPITANT"
    ],
    };
    }

  // Acute Gastroenteritis | Vomiting, Rads Declined
    function generateCanineAcuteGastroenteritisVomitingRadsDeclinedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["ACUTE_GASTROENTERITIS_VOMITING_ONLY"],
    text: [
    `Acute gastroenteritis: At this time no overt causes of vomiting were identified. X-rays can be performed to differentiate between something being stuck in the intestines and dietary indiscretion (eating something that isn’t healthy for dogs). You’ve declined x-rays in favour of symptomatic treatment at this time. Give the medication as prescribed below. If you still see vomiting within 24 hours of the injection, your dog needs to go to your nearest veterinary emergency hospital immediately.`
    ].join('\n'),

    boldKeys: [
      "ACUTE_GASTROENTERITIS_HEADER"
    ],

    boldUnderlineKeys: [
      "VOMITING_POST_MAROPITANT"
    ],
    };
    }

  // Anal Glands | Full, Expressed
    function generateCanineAnalGlands1FullExpressedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["FULL_ANAL_GLANDS"],
    text: [
    `Anal glands: Your dog had full anal glands that were expressed at the clinic. Dogs have anal glands on either side of their anus to leave their scent on their stool. Typically this empties whenever they defecate, but dogs with soft stool or diarrhea have difficulty expressing them. You can add psyllium husk to increase the fiber content if stools are soft or watery. Some dogs have anal glands that never empty correctly. If your dog’s anal glands are full, you may see scooting on the floor or over fixation on the anus.`,
      `You can learn how to express your dog’s anal glands yourself by reading the Emptying a Dog or Cat's Anal Sacs article on Veterinary Partner. Otherwise you can visit a clinic or groomer to have them expressed. If your dog continues to fixate on the anus, bring back a stool sample & we can test for intestinal parasites.`
    ].join('\n'),

    boldKeys: [
      "ANAL_GLANDS_HEADER"
    ],

    boldUnderlineKeys: [
      "ANAL_GLANDS_EXPRESSED",
      "ANAL_GLANDS_BRING_STOOL",
      "ANAL_GLAND_SYMPTOMS",
    ],

    linkKeys: [
      "EMPTYING_A_DOG_OR_CAT_S_ANAL_SACS_ARTICLE"
    ],
    };
    }

  // Anal Glands | 2nd, Known
    function generateCanineAnalGlands2KnownTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["FULL_ANAL_GLANDS"],
    text: [
    `Anal glands: Your dog had full anal glands that were expressed at the clinic. If you’d like you to learn how to express them yourself, you can learn more from the Emptying a Dog or Cat's Anal Sacs article on Veterinary Partner. Otherwise bring your pet back as needed. If your dog continues to scoot on the floor, bring back a stool sample & we can test for intestinal parasites.`
    ].join('\n'),

    boldKeys: [
      "ANAL_GLANDS_HEADER"
    ],

    boldUnderlineKeys: [
      "ANAL_GLANDS_EXPRESSED",
      "ANAL_GLANDS_STILL_SCOOTING"
    ],

    linkKeys: [
      "EMPTYING_A_DOG_OR_CAT_S_ANAL_SACS_ARTICLE"
    ],
    };
    }

  // Anal Glands | 3rd, Infected
    function generateCanineAnalGlands3InfectedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    return {
    sex,
    plurality,
    diagnoses: ["INFECTED_ANAL_GLANDS"],
    text: [
    `Infected anal glands: Your dog had full anal glands that appeared infected when they were expressed in the clinic. Dogs have anal glands on either side of their anus to leave their scent on their stool. Typically this empties whenever they defecate, but dogs with soft stool or diarrhea have difficulty expressing them.  If your dog’s anal glands are full, you may see scooting on the floor or over fixation on the anus. Ultimately this can lead to infection, similar to what your dog appears to have today. You can add psyllium husk to increase the fiber content if stools are soft or watery.`,
      `Your dog has been administered antibiotics and pain control for the infection. You can learn how to express your dog’s anal glands yourself by reading the Emptying a Dog or Cat's Anal Sacs article on Veterinary Partner. Otherwise you can visit a clinic or groomer to have them expressed. If your dog continues to fixate on the anus, bring back a stool sample & we can test for intestinal parasites.`
    ].join('\n'),

    boldKeys: [
      "INFECTED_ANAL_GLANDS_HEADER"
    ],

    boldUnderlineKeys: [
      "ANAL_GLANDS_BRING_STOOL",
      "ANAL_GLAND_SYMPTOMS",
      "INFECTED_ANAL_GLANDS_SEEN"
    ],

    linkKeys: [
      "EMPTYING_A_DOG_OR_CAT_S_ANAL_SACS_ARTICLE"
    ],
    };
    }

  // Periodontal Disease | Mild
    function generateCanine1PeriodontalDiseaseTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Periodontal disease: Your ${g.dog} ${g.has} early dental disease. While brushing ${g.his} teeth is the best way to keep them clean, it  will not remove the tartar & calculus that is already there. Schedule a dental cleaning to completely remove calculus and then try brushing the teeth.`,
    `Until then, use a small dog toothbrush, medium/large dog toothbrush & animal safe toothpaste such as C.E.T. Start by having ${g.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then apply the pet safe toothpaste & let ${g.him} lick it off every day for a week. Finally, gently brush ${g.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
    `If your ${g.dog} ${g.resists} having ${g.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${g.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`
    ].join('\n');

    return {
    sex,
    plurality,
    text,
    diagnoses: ["MILD_PERIODONTAL_DISEASE"],
    cleanupKeys: ["DENTAL_HEADER"],
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

  // Periodontal Disease | Moderate
    function generateCanine2PeriodontalDiseaseTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Periodontal disease: Your ${g.dog} ${g.needs} a Complete Oral Health Assessment and Treatment (COHAT) procedure. While brushing ${g.his} teeth is the best way to keep them clean, it will not remove the tartar & calculus that is already there. Schedule a dental cleaning within the next three months.`,
    `Brushing can still be performed right now but will be most effective after the next cleaning. Wait 3 weeks after ${g.his} next teeth cleaning before brushing the teeth to allow time for the mouth’s soreness to abate. Use a small dog toothbrush, canine toothbrush, or finger toothbrush & animal safe toothpaste such as C.E.T. Start by having ${g.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then apply the pet safe toothpaste & let ${g.him} lick it off every day for a week. Finally, gently brush ${g.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
    `If your ${g.dog} ${g.resists} having ${g.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${g.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`
    ].join('\n');

    return {
    sex,
    plurality,
    text,
    diagnoses: ["MODERATE_PERIODONTAL_DISEASE"],
    cleanupKeys: ["DENTAL_HEADER"],
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

  // Periodontal Disease | Severe
    function generateCanine3PeriodontalDiseaseTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Periodontal disease: Your ${g.dog} ${g.needs} a Complete Oral Health Assessment and Treatment (COHAT) procedure. Brushing ${g.his} teeth is the best way to keep them clean, but it will not remove the tartar & calculus that is already there. In fact, brushing right now is not advised as it will likely cause ${g.him} pain and possibly bleeding given how severe the disease is. Schedule a dental cleaning within the next three months.`,
    `In the meantime, feed soft food such as wet food or dry food soaked in a few tablespoons of warm water 30 seconds prior to feeding. This will make it easier for your dog to chew. Wait 3 weeks after ${g.his} next teeth cleaning before brushing the teeth to allow time for the mouth’s soreness to abate. Use a small dog toothbrush, medium/large dog toothbrush, or finger toothbrush & animal safe toothpaste such as C.E.T. Start by having ${g.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then apply the pet safe toothpaste & let ${g.him} lick it off every day for a week. Finally, gently brush ${g.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
    `If your ${g.dog} ${g.resists} having ${g.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${g.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`
    ].join('\n');

    return {
    sex,
    plurality,
    text,
    diagnoses: ["SEVERE_PERIODONTAL_DISEASE"],
    cleanupKeys: ["DENTAL_HEADER"],
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

  // Periodontal Disease | Age Restricted
    function generateCanine4PeriodontalDiseaseAgeTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Periodontal disease: Your ${g.dog} ${g.shows} signs of dental disease. However, older patients are more at risk of anesthesia complications. A routine dental cleaning is not advised in your ${g.dog} for that reason unless it is performed with a dental specialist. The recommended clinic is Veterinary Dental Specialists. A referral to those clinics can be facilitated at your request.`,
    `In the meantime, feed soft food such as wet food or dry food soaked in a few tablespoons of warm water 30 seconds prior to feeding. This will make it easier for your ${g.dog} to chew the food. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`
    ].join('\n');

    return {
    sex,
    plurality,
    text,
    diagnoses: ["PERIODONTAL_DISEASE"],
    cleanupKeys: ["DENTAL_HEADER"],
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

  // Periodontal Disease | Concurrent Disease
    function generateCanine4PeriodontalDiseaseConcurrentDiseaseTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Periodontal disease: Your ${g.dog} ${g.shows} signs of dental disease. However, a dental cleaning is not recommended at this time until the current problem is dealt with. Once the other problem has been addressed, a dental cleaning can be performed. In the meantime, feed soft food such as wet food or dry food soaked in a few tablespoons of warm water 30 seconds prior to feeding. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`
    ].join('\n');

    return {
    sex,
    plurality,
    text,
    diagnoses: ["PERIODONTAL_DISEASE"],
    cleanupKeys: ["DENTAL_HEADER"],
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

  // Periodontal Disease | Heart Murmur
    function generateCanine4PeriodontalDiseaseHeartMurmurTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Periodontal disease: Your ${g.dog} ${g.shows} signs of dental disease. However, a dental cleaning is not recommended at this time until the heart is further investigated. Once the other problem has been addressed, a dental cleaning can be performed. In the meantime, feed soft food such as wet food or dry food soaked in a few tablespoons of warm water 30 seconds prior to feeding. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`
    ].join('\n');

    return {
    sex,
    plurality,
    text,
    diagnoses: ["PERIODONTAL_DISEASE"],
    cleanupKeys: ["DENTAL_HEADER"],
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
  // Atopic Dermatitis | Antihistamines 1
    function generateCanineAtopicDermatitisMild1Template(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);

    const text = [
    `Atopic dermatitis: Unlike humans where allergies present in the respiratory tract (runny nose, sneezing/coughing, etc.), allergies in pets usually appear in the skin (shaking the head, chewing/licking the paws, scratching excessively, etc.). In fact, one of the most common causes of chronic ear infections is allergies.`,
    `At this time we will not be starting with daily oral medicine (Apoquel and Zenrelia) or monthly injectable medicine (Cytopoint). You can give over the counter antihistamines such as Benadryl 25mg (give up to 1 tablet per 25 lbs every 12 hours) or Zyrtec 10mg (give up to 1 tablet per 10 lbs every 12 - 24 hours). Potential side effects (drowsiness, increased drinking) are more common with Benadryl than Zyrtec.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Atopic Dermatitis | Antihistamines 2
    function generateCanineAtopicDermatitisMild2Template(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies which you currently give over the counter antihistamines for. Continue to give Benadryl 25mg (give up to 1 tablet per 25 lbs every 12 hours) or Zyrtec 10mg (give up to 1 tablet per 10 lbs every 12 - 24 hours) as needed. If you feel like allergies are not well controlled, prescription medicine such as Cytopoint (an injection given every 4 - 8 weeks) or Apoquel OR Zenrelia (oral pills given every 24 hours) can be given for better control.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Atopic Dermatitis | Apoquel 1
    function generateCanineAtopicDermatitis1ApoquelTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Atopic dermatitis: Unlike humans where allergies presents in the respiratory tract (runny nose, sneezing, etc.), allergies in pets usually appears in the skin (shaking the head, chewing/licking the paws, scratching excessively, etc.). In fact, one of the most common causes of chronic ear infections is allergies. While antihistamines (Benadryl, Zyrtec, etc.) occasionally help, your dog shows signs of severe allergies.`,
    `Apoquel has been sent home to resolve allergies. Give as prescribed. If itching & scratching persists after two weeks, Cytopoint (an injection given every 4 - 8 weeks) can be tried instead. Alternatively, they can be given together to have a more powerful effect to control allergies. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`,
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Atopic Dermatitis | Apoquel 2, Maintenance
    function generateCanineAtopicDermatitis2ApoquelTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies & gets Apoquel to control them. If itching & scratching persists, Cytopoint (an injection given every 4 - 8 weeks) can be tried instead. Alternatively, they can be given together to have a more powerful effect to control allergies. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Atopic Dermatitis | Apoquel 3, Add Cytopoint
    function generateCanineAtopicDermatitis3ApoquelTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies & gets Apoquel to control them. However, Apoquel on its own doesn’t appear effective enough to control allergies. As such we will be adding Cytopoint to the plan. These medications improve the effectiveness of the other and are safe to give together.`,
    `Continue to give Apoquel as you’ve been doing. If you see full allergy control, you can try discontinuing Apoquel in 2 weeks to see if Cytopoint on its own can help control allergies. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Atopic Dermatitis | Apoquel 4 Switch to Zenrelia
    function generateCanineAtopicDermatitis4ApoquelTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies & gets Apoquel to control them. However, Apoquel doesn’t appear to be effective enough. We will be switching your dog to Zenrelia instead to see if this better controls allergies. Give daily for 1 month for best results. Do not give Zenrelia in the same 24 hours as Apoquel.`,
    `If itching & scratching persists, Cytopoint (an injection given every 4 - 8 weeks) can be tried instead. Alternatively, they can be given together to have a more powerful effect to control allergies. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Atopic Dermatitis | Cytopoint 1
    function generateCanineAtopicDermatitis1CytopointTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Atopic dermatitis: Unlike humans where allergies presents in the respiratory tract (runny nose, sneezing, etc.), allergies in pets usually appears in the skin (shaking the head, chewing/licking the paws, scratching excessively. etc.). In fact, one of the most common causes of chronic ear infections is allergies. While antihistamines (Benadryl, Zyrtec, etc.) occasionally help, your dog shows signs of severe allergies.`,
    `Cytopoint has been given in clinic to resolve allergies. It typically lasts 4 - 8 weeks. If itching & scratching occurs before 4 weeks, Apoquel OR Zenrelia (oral pills given once a day) can be sent home in addition to monthly Cytopoint injections to better control allergies.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Atopic Dermatitis | Cytopoint 2
    function generateCanineAtopicDermatitis2CytopointTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies & gets Cytopoint injections to control them. The injection was given today & typically lasts 4 - 8 weeks. If the allergies return before 4 weeks, your dog may need Apoquel or Zenrelia in addition to Cytopoint. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`
    ].join('\n');

    return {
    sex,
    plurality,
    text,
    diagnoses: ["ATOPIC_DERMATITIS"],
    boldKeys: [
    'ATOPIC_DERMATITIS_HEADER',
    'ANTIHISTAMINE_ADDITION',
    ],

    boldUnderlineKeys: [
    `CYTOPOINT_ADDITIONAL_SUPPORT`,
    ],
    };
    }

  // Atopic Dermatitis | Meds Declined
    function generateCanineAtopicDermatitisMedsDeclinedTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Atopic dermatitis: Unlike humans where allergies presents in the respiratory tract (runny nose, sneezing, etc.), allergies in pets usually appears in the skin (shaking the head, chewing/licking the paws, scratching excessively. etc.). In fact, one of the most common causes of chronic ear infections is allergies.`,
    `Cytopoint (an injection given every 4 - 8 weeks) or either Apoquel OR Zenrelia (oral pills given every 24 hours) are more effective than over the counter medicine and are advised, but you have elected to try antihistamines first. You can give over the counter antihistamines such as Benadryl 25mg (give up to 1 tablet per 25 lbs every 12 hours) or Zyrtec 10mg (give up to 1 tablet per 10 lbs every 12 - 24 hours). Side effects (drowsiness, increased drinking) are more common with Benadryl than Zyrtec.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Atopic Dermatitis | Zenrelia 1
    function generateCanineAtopicDermatitis1ZenreliaTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Atopic dermatitis: Unlike humans where allergies presents in the respiratory tract (runny nose, sneezing, etc.), allergies in pets usually appears in the skin (shaking the head, chewing/licking the paws, scratching excessively, etc.). In fact, one of the most common causes of chronic ear infections is allergies. While antihistamines (Benadryl, Zyrtec, etc.) occasionally help, your dog shows signs of severe allergies.`,
    `Zenrelia has been sent home to resolve allergies. Give as prescribed. If itching & scratching persists after two weeks, Cytopoint (an injection given every 4 - 8 weeks) can be tried instead. Alternatively, they can be given together to have a more powerful effect to control allergies.`,
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Atopic Dermatitis | Zenrelia 2, Maintenance
    function generateCanineAtopicDermatitis2ZenreliaTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies & gets Zenrelia to control them. If itching & scratching persists, Cytopoint (an injection given every 4 - 8 weeks) can be tried instead. Alternatively, they can be given together to have a more powerful effect to control allergies. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`,
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Atopic Dermatitis | Zenrelia 3, Add Cytopoint
    function generateCanineAtopicDermatitis3ZenreliaTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Atopic dermatitis: Your dog is known to have allergies & gets Zenrelia to control them. However, Zenrelia on its own doesn’t appear effective enough to control allergies. As such we will be adding Cytopoint to the plan. These medications improve the effectiveness of the other and are safe to give together.`,
    `Continue to give Zenrelia as you’ve been doing. If you see full allergy control, you can try discontinuing Zenrelia in 2 weeks to see if Cytopoint on its own can help control allergies. You can still give Benadryl 25mg (1 tablet per 25 lbs every 12 hours) or Zyrtec (up to 1 tablet per 10 lbs every 12 - 24 hours) for additional support.`
    ].join('\n');

    return {
    sex,
    plurality,
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
  // Osteoarthritis | 1st NSAID, Initial
    function generateCanineOsteoarthritis1NSAIDTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Osteoarthritis: Arthritis was detected in your dog’s joints. The most common sign of this is being slow after waking up/laying down for a while or being sore after walks. There are three options for treatment: monthly injectable medicine, twice daily oral pain medicine, and joint supplements. You’ve elected to try a non-steroidal anti-inflammatory drug (NSAID) to reduce pain & inflammation from arthritis. Bloodwork is required every 6 - 12 months while on NSAIDs.`,
    `Alternatives include Librela (an injection given in clinic once a month), gabapentin (oral capsules), & joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size). Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Osteoarthritis | 2nd NSAID, Maintenance
    function generateCanineOsteoarthritis2NSAIDTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Osteoarthritis: Your dog is known to have arthritis & is currently on a non-steroidal anti-inflammatory drug (NSAID) to reduce pain & inflammation from arthritis. Bloodwork is required every 6 - 12 months while on NSAIDs. Alternatives include Librela (an injection given in clinic once a month), gabapentin (oral capsules), & joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size). Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Osteoarthritis | 3rd NSAID, Switch NSAIDs
    function generateCanineOsteoarthritis3NSAIDTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Osteoarthritis: Your dog is known to have arthritis & is currently on a non-steroidal anti-inflammatory drug (NSAID) to reduce pain & inflammation from arthritis. We will be switching to a different NSAID to see if better control is provided while still being safe for your pet. Bloodwork is required every 6 - 12 months while on NSAIDs.`,
    `If we still don’t see the improvement we’d like, we can discuss adding on alternatives such as Librela (an injection given in clinic once a month), gabapentin (oral capsules), & joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size). Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Osteoarthritis | 1st Gabapentin
    function generateCanineOsteoarthritis1GabapentinTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Osteoarthritis: Arthritis was detected in your dog’s joints. The most common sign of this is being slow after waking up/laying down for a while or being sore after walks. There are three options for treatment: monthly injectable medicine, twice daily oral pain medicine, and joint supplements.`,
    `Librela is an injection given once a month that controls arthritis in most patients. However, it may take 2 - 3 months before improvement is seen. Instead, immediate relief can be provided via NSAIDs such as carprofen & grapiprant. Bloodwork is recommended every 6 - 12 months while on NSAIDs. `,
    `You’ve elected to try gabapentin. While less effective than NSAIDs, they do not require bloodwork and still offer excellent pain control. Joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size) are a more natural alternative that you can add on to improve mobility, but they do not reduce pain on their own. Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Osteoarthritis | 2nd Gabapentin, Continue
    function generateCanineOsteoarthritis2GabapentinTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Osteoarthritis: Your dog is known to have arthritis & is currently on gabapentin. Alternatives to gabapentin include Librela, an injection that can be given once a month with minimal side effects. However, it may take 2 - 3 months before improvement is seen. Non-steroidal anti-inflammatory drugs such as carprofen & grapiprant can also control arthritis so long as steroids aren’t currently being given. Bloodwork is recommended every 6 - 12 months while on NSAIDs.`,
    `Joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size) can also be added to any treatment plan to increase mobility but will not control pain. Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Osteoarthritis | 1st Joint Supplements
    function generateCanineOsteoarthritis1JointSupplementsTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Osteoarthritis: Arthritis was detected in your dog’s joints. The most common sign of this is being slow after waking up/laying down for a while or being sore after walks. There are three options for treatment: monthly injectable medicine, twice daily oral pain medicine, and joint supplements. Immediate relief can be provided via gabapentin or non-steroidal anti-inflammatory drugs (NSAIDs) such as carprofen & grapiprant. NSAIDs are more effective at controlling arthritis than gabapentin & require bloodwork every 6 months. Librela is an injection given once a month that controls arthritis in most patients. However, it may take 2 - 3 months before improvement is seen.`,
    `At this time you’ve elected to try joint supplements. Joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size) are a more natural alternative to increase mobility but do not reduce pain. Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Osteoarthritis | 2nd Joint Supplements, Continue
    function generateCanineOsteoarthritis2JointSupplementsTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Osteoarthritis: Your dog is known to have arthritis & currently gets joint supplements. Joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size) are a more natural alternative to increase mobility but do not reduce pain. Non-steroidal anti-inflammatory drugs (NSAIDs) such as carprofen & grapiprant provide immediate relief from arthritis so long as steroids aren’t currently being given. Bloodwork is required every 6 months while on NSAIDs. `,
    `Gabapentin can be used in addition to or instead of NSAIDs with minimal side effects, though it isn’t as effective as NSAIDs. Finally, the Librela injection can be given once a month though it can take 2 - 3 months to see improvement. If you have concerns about your dog’s arthritis, contact the clinic & we can discuss which medicine is best for your dog. Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Osteoarthritis | 1st Librela
    function generateCanineOsteoarthritis1LibrelaTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Osteoarthritis: Arthritis was detected in your dog’s joints. The most common sign of this is being slow after waking up/laying down for a while or being sore after walks. There are three options for treatment: monthly injectable medicine, twice daily oral pain medicine, and joint supplements.`,
    `Librela, an injection given once a month that controls arthritis, was given today. Watch for signs of reaction including vomiting, diarrhea, lethargy, & excessive panting/fever. If signs are seen, bring your dog back immediately as these are signs of a reaction. Librela may take 2 - 3 months before full effects are seen, so oral medicine (gabapentin, carprofen, grapiprant, etc.) can be used in the meantime if necessary.`,
    `You can continue giving joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size) to help improve joint mobility. Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Osteoarthritis | 2nd Librela
    function generateCanineOsteoarthritis2LibrelaTemplate(sex, plurality = 'singular') {
    const g = getGrammar('wellness', plurality, sex);
    const text = [
    `Osteoarthritis: Your dog is known to have arthritis & currently gets Librela injections. Watch for signs of reaction including vomiting, diarrhea, lethargy, & excessive panting/fever. Bring your dog back immediately if any are seen. Joint supplements (aim for those with glucosamine and at least 1,000mg of DHA & EPA per serving size) can be added in addition to further improve mobility. Keeping your dog an appropriate weight can also help reduce joint pain.`
    ].join('\n');

    return {
    sex,
    plurality,
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

  // Puppy Wellness Definitions
    '/cReset': () => generateCanineResetTemplate(),
    '/fReset': () => generateFelineResetTemplate(),
    '/c8wksSmall': (sex, plurality) => generate8WkWellnessTemplate('small', sex, plurality),
    '/c8wksLarge': (sex, plurality) => generate8WkWellnessTemplate('large', sex, plurality),

    '/c12wksSmall': (sex, plurality) => generate12WkWellnessTemplate('small', sex, plurality),
    '/c12wksLarge': (sex, plurality) => generate12WkWellnessTemplate('large', sex, plurality),

    '/c16wksSmall': (sex, plurality) => generate16WkWellnessTemplate('small', sex, plurality),
    '/c16wksLarge': (sex, plurality) => generate16WkWellnessTemplate('large', sex, plurality),

  // Canine Adult Wellness Definitions
    '/cInitialAdult': (sex, plurality) => generateInitialAdultTemplate(sex, plurality),
    '/c1year': (sex, plurality) => generate1YearAdultTemplate(sex, plurality),
    '/c2year': (sex, plurality) => generate2YearAdultTemplate(sex, plurality),
    '/c2yearLepto': (sex, plurality) => generate2YearLeptoTemplate(sex, plurality),
    '/c7year': (sex, plurality) => generate7YearAdultTemplate(sex, plurality),
    '/c7yearLepto': (sex, plurality) => generate7YearLeptoTemplate(sex, plurality),

    '/cOverweight1': (sex, plurality) => generateCanineOverweightTemplate(sex, plurality),
    '/cOverweight2': (sex, plurality) => generateCanineOverweight2Template(sex, plurality),
    '/cHealthyWeight': (sex, plurality) => generateCanineHealthyWeightTemplate(sex, plurality),
    '/cUnderweight': (sex, plurality) => generateCanineUnderweightTemplate(sex, plurality),

  // Canine Ophthalmology Definitions
    '/cBlind0Partial': (sex, plurality) => generateCanineBlind0PartialTemplate(sex, plurality),
    '/cBlind1': (sex, plurality) => generateCanineBlind1Template(sex, plurality),
    '/cBlind2Known': (sex, plurality) => generateCanineBlind2KnownTemplate(sex, plurality),
    '/cCherryEye': (sex, plurality) => generateCanineCherryEyeTemplate(sex, plurality),
    '/cCherryEyes': (sex) => generateCanineCherryEyeTemplate(sex, "plural"),
    '/cCompleteCataract': (sex, plurality) => generateCompleteCataractsTemplate(sex, plurality),
    '/cCompleteCataracts': (sex) => generateCompleteCataractsTemplate(sex, "plural"),
    '/cConjunctivitisTestsDeclined': (sex, plurality) => generateCanineConjunctivitisTestsDeclinedTemplate(sex, plurality),
    '/cConjunctivitisDiagnosed': (sex, plurality) => generateCanineConjunctivitisDiagnosedTemplate(sex, plurality),
    '/cCornealUlcer': (sex) => generateCanineCornealUlcerTemplate(sex, "singular"),
    '/cCornealUlcers': (sex) => generateCanineCornealUlcerTemplate(sex, "plural"),
    '/cEntropion': (sex, plurality) => generateCanineEntropionTemplate(sex, plurality),
    '/cGlaucoma1': (sex, plurality) => generateCanineGlaucoma1DiagnosedTemplate(sex, plurality),
    '/cKeratoconjunctivitisSicca1Diagnosed': (sex, plurality) => generateCanineKeratoconjunctivitisSicca1DiagnosedTemplate(sex, plurality),
    '/cKeratoconjunctivitisSicca2Controlled': (sex, plurality) => generateCanineKeratoconjunctivitisSicca2ControlledTemplate(sex, plurality),
    '/cMeibomianGlandAdenoma1': (sex, plurality) => generateCanineMeibomianGlandAdenoma1PresumedTemplate(sex, plurality),
    '/cNuclearSclerosis': (sex, plurality) => generateCanineNuclearSclerosisTemplate(sex, plurality),

  // Canine Cardiology Definitions
    '/c2ndDegreeAVBlock': (sex, plurality) => generateCanine2ndDegreeAVBlockTemplate(sex, plurality),
    '/cHeartMurmur0': (sex, plurality) => generateCanineHeartMurmur0Template(sex, plurality),
    '/cHeartMurmur1RadiographsNormal': (sex, plurality) => generateCanineHeartMurmur1RadiographsNormalTemplate(sex, plurality),
    '/cHeartMurmur1Cardiomegaly': (sex, plurality) => generateCanineHeartMurmur1CardiomegalyTemplate(sex, plurality),
    '/cHeartMurmur3Known': (sex, plurality) => generateCanineHeartMurmur3KnownTemplate(sex, plurality),
    '/cHeartwormsAdulticidalTreatment': (sex, plurality) => generateCanineHeartwormsAdulticidalTreatmentTemplate(sex, plurality),
    '/cHeartwormsSlowKillHealthyTreatment': (sex, plurality) => generateCanineHeartwormsSlowKillHealthyDogTreatmentTemplate(sex, plurality),
    '/cHeartwormTestRepeat': (sex, plurality) => generateCanineHeartwormTestRepeatTemplate(sex, plurality),
    '/cHypertension1Diagnosed': (sex, plurality) => generateCanineHypertension1DiagnosedTemplate(sex, plurality),
    '/cLeftSidedCongestiveHeartFailure': (sex, plurality) => generateCanineLeftSidedCongestiveHeartFailureTemplate(sex, plurality),
    '/cMyxomatousMitralValveDisease': (sex, plurality) => generateCanineMyxomatousMitralValveDiseaseTemplate(sex, plurality),

  // Canine Respiratory Definitions
    '/cBordetellosis0Presumed': (sex, plurality) => generateCanineBordetellosis0PresumedTemplate(sex, plurality),
    '/cBrachycephalicObstructiveAirwaySyndrome': (sex, plurality) => generateCanineBrachycephalicObstructiveAirwaySyndromeTemplate(sex, plurality),
    '/cChronicBronchitis0Presumed': (sex, plurality) => generateCanineChronicBronchitis0PresumedTemplate(sex, plurality),
    '/cCollapsingTrachea1Theophylline': (sex, plurality) => generateCanineCollapsingTrachea1TheophyllineTemplate(sex, plurality),
    '/cCollapsingTrachea3NoMeds': (sex, plurality) => generateCanineCollapsingTrachea3NoMedsTemplate(sex, plurality),
    '/cLaryngealParalysis1Diagnosed': (sex, plurality) => generateCanineLaryngealParalysis1DiagnosedTemplate(sex, plurality),
    '/cReverseSneezing': (sex, plurality) => generateCanineReverseSneezingTemplate(sex, plurality),

  // Canine Endocrine Definitions
    '/cDiabetesMellitus1Diagnosed': (sex, plurality) => generateCanineDiabetesMellitus1DiagnosedTemplate(sex, plurality),
    '/cDiabetesMellitus3Controlled': (sex, plurality) => generateCanineDiabetesMellitus3ControlledTemplate(sex, plurality),
    '/cHypothyroidism1Diagnosed': (sex, plurality) => generateCanineHypothyroidism1DiagnosedTemplate(sex, plurality),
    '/cHypothyroidism2Recheck': (sex, plurality) => generateCanineHypothyroidism2RecheckTemplate(sex, plurality),
    '/cHypothyroidism3Controlled': (sex, plurality) => generateCanineHypothyroidism3ControlledTemplate(sex, plurality),
    '/cHyperadrenocorticism1ACTHStimTest': (sex, plurality) => generateCanineHyperadrenocorticism1ACTHStimTestTemplate(sex, plurality),
    '/cHyperadrenocorticism2Diagnosed': (sex, plurality) => generateCanineHyperadrenocorticism2DiagnosedTemplate(sex, plurality),
    '/cHyperadrenocorticism3Controlled': (sex, plurality) => generateCanineHyperadrenocorticism3ControlledTemplate(sex, plurality),
    '/cHyperadrenocorticism4Uncontrolled': (sex, plurality) => generateCanineHyperadrenocorticism4UncontrolledTemplate(sex, plurality),
    '/cHyperadrenocorticism5Checkup': (sex, plurality) => generateCanineHyperadrenocorticism5CheckupTemplate(sex, plurality),
    '/cPancreatitis1Diagnosed': (sex, plurality) => generateCaninePancreatitis1DiagnosedTemplate(sex, plurality),

  // Gastrointestinal Definitions
    '/cAcuteGastroenteritisDiarrheaHomeDiet': (sex, plurality) => generateCanineAcuteGastroenteritisDiarrheaHomeDietTemplate(sex, plurality),
    '/cAcuteGastroenteritisDiarrheaFecalTest': (sex, plurality) => generateCanineAcuteGastroenteritisDiarrheaFecalTestTemplate(sex, plurality),
    '/cAcuteGastroenteritisDiarrheaDeclinedFecalTest': (sex, plurality) => generateCanineAcuteGastroenteritisDiarrheaDeclinedFecalTestTemplate(sex, plurality),
    '/cAcuteGastroenteritisVomitingDiarrheaBloodworkFecalTest': (sex, plurality) => generateCanineAGDiarrheaVomitingBloodworkFecalTestTemplate(sex, plurality),
    '/cAcuteGastroenteritisVomitingBloodworkNormal': (sex, plurality) => generateCanineAcuteGastroenteritisVomitingBloodworkNormalTemplate(sex, plurality),
    '/cAcuteGastroenteritisVomitingBloodworkDeclined': (sex, plurality) => generateCanineAcuteGastroenteritisVomitingBloodworkDeclinedTemplate(sex, plurality),
    '/cAcuteGastroenteritisVomitingRadiographsNormal': (sex, plurality) => generateCanineAcuteGastroenteritisVomitingRadiographsNormalTemplate(sex, plurality),
    '/cAcuteGastroenteritisVomitingRadiographsDeclined': (sex, plurality) => generateCanineAcuteGastroenteritisVomitingRadsDeclinedTemplate(sex, plurality),
    '/cAnalGlands1FullExpressed': (sex, plurality) => generateCanineAnalGlands1FullExpressedTemplate(sex, plurality),
    '/cAnalGlands2Known': (sex, plurality) => generateCanineAnalGlands2KnownTemplate(sex, plurality),
    '/cAnalGlands3Infected': (sex, plurality) => generateCanineAnalGlands3InfectedTemplate(sex, plurality),
    '/cPeriodontalDisease1': (sex, plurality) => generateCanine1PeriodontalDiseaseTemplate(sex, plurality),
    '/cPeriodontalDisease2': (sex, plurality) => generateCanine2PeriodontalDiseaseTemplate(sex, plurality),
    '/cPeriodontalDisease3': (sex, plurality) => generateCanine3PeriodontalDiseaseTemplate(sex, plurality),
    '/cPeriodontalDisease4Age': (sex, plurality) => generateCanine4PeriodontalDiseaseAgeTemplate(sex, plurality),
    '/cPeriodontalDisease4ConcurrentDisease': (sex, plurality) => generateCanine4PeriodontalDiseaseConcurrentDiseaseTemplate(sex, plurality),
    '/cPeriodontalDisease4HeartMurmur': (sex, plurality) => generateCanine4PeriodontalDiseaseHeartMurmurTemplate(sex, plurality),

  // Musculoskeletal Definitions
    '/cOsteoarthritis1NSAID': (sex, plurality) => generateCanineOsteoarthritis1NSAIDTemplate(sex, plurality),
    '/cOsteoarthritis2NSAID': (sex, plurality) => generateCanineOsteoarthritis2NSAIDTemplate(sex, plurality),
    '/cOsteoarthritis3NSAID': (sex, plurality) => generateCanineOsteoarthritis3NSAIDTemplate(sex, plurality),
    '/cOsteoarthritis1Gabapentin': (sex, plurality) => generateCanineOsteoarthritis1GabapentinTemplate(sex, plurality),
    '/cOsteoarthritis2Gabapentin': (sex, plurality) => generateCanineOsteoarthritis2GabapentinTemplate(sex, plurality),
    '/cOsteoarthritis1JointSupplements': (sex, plurality) => generateCanineOsteoarthritis1JointSupplementsTemplate(sex, plurality),
    '/cOsteoarthritis2JointSupplements': (sex, plurality) => generateCanineOsteoarthritis2JointSupplementsTemplate(sex, plurality),
    '/cOsteoarthritis1Librela': (sex, plurality) => generateCanineOsteoarthritis1LibrelaTemplate(sex, plurality),
    '/cOsteoarthritis2Librela': (sex, plurality) => generateCanineOsteoarthritis2LibrelaTemplate(sex, plurality),

  // Immunology Definitions
    '/cVaccineInformation': (sex, plurality) => generateCanineVaccineInformationTemplate(sex, plurality),

  // Dermatology/ Definitions
    '/cAtopicDermatitis1Antihistamines': (sex, plurality) => generateCanineAtopicDermatitisMild1Template(sex, plurality),
    '/cAtopicDermatitis2Antihistamines': (sex, plurality) => generateCanineAtopicDermatitisMild2Template(sex, plurality),
    '/cAtopicDermatitis1Apoquel': (sex, plurality) => generateCanineAtopicDermatitis1ApoquelTemplate(sex, plurality),
    '/cAtopicDermatitis2Apoquel': (sex, plurality) => generateCanineAtopicDermatitis2ApoquelTemplate(sex, plurality),
    '/cAtopicDermatitis3Apoquel': (sex, plurality) => generateCanineAtopicDermatitis3ApoquelTemplate(sex, plurality),
    '/cAtopicDermatitis4Apoquel': (sex, plurality) => generateCanineAtopicDermatitis4ApoquelTemplate(sex, plurality),
    '/cAtopicDermatitis1Cytopoint': (sex, plurality) => generateCanineAtopicDermatitis1CytopointTemplate(sex, plurality),
    '/cAtopicDermatitis2Cytopoint': (sex, plurality) => generateCanineAtopicDermatitis2CytopointTemplate(sex, plurality),
    '/cAtopicDermatitisMedsDeclined': (sex, plurality) => generateCanineAtopicDermatitisMedsDeclinedTemplate(sex, plurality),
    '/cAtopicDermatitis1Zenrelia': (sex, plurality) => generateCanineAtopicDermatitis1ZenreliaTemplate(sex, plurality),
    '/cAtopicDermatitis2Zenrelia': (sex, plurality) => generateCanineAtopicDermatitis2ZenreliaTemplate(sex, plurality),
    '/cAtopicDermatitis3Zenrelia': (sex, plurality) => generateCanineAtopicDermatitis3ZenreliaTemplate(sex, plurality),
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

/* ------------------ EXPAND KEYWORDS ENGINE ------------------ */
  // Main Function
    function runExpansionEngine(matches) {
    const body = DocumentApp.getActiveDocument().getBody();
    const cleanupQueue = [];

    // 1. PRIORITY PASS: Handle Resets First
    const resetMatch = matches.find(m => m.normalized.endsWith('reset'));
    if (resetMatch) {
    const { base, sex, plurality } = parseKeywordMetadata(resetMatch.normalized);
    const templateFn = TEMPLATE_DEFINITIONS[base];
    if (templateFn) {
    const template = templateFn(sex, plurality);
    body.clear(); 
    insertTemplateAtIndex(body, template, 0);
    if (template.customAction) template.customAction();
    }
    }

    // 2. SECOND PASS: Process medications and buffer standard templates
    matches.forEach(m => {
    if (m.normalized.endsWith('reset')) return; 

    const { base, sex, plurality } = parseKeywordMetadata(m.normalized);
    const templateFn = TEMPLATE_DEFINITIONS[base];

    if (templateFn) {
    const template = templateFn(sex, plurality);

    if (template.cleanupKeys && template.cleanupKeys.length > 0) {
    cleanupQueue.push(...template.cleanupKeys);
    }

    let rank = template.rank || 999;
    if (template.diagnoses && template.diagnoses.length > 0) {
    bufferDiagnoses(template.diagnoses);
    const firstKey = template.diagnoses[0];
    const diagRank = (DIAGNOSIS_REGISTRY[firstKey] && DIAGNOSIS_REGISTRY[firstKey].rank) || 999;
    rank = Math.min(rank, diagRank);
    }

    bufferTemplate(template, rank);
    if (template.customAction) template.customAction();
    }

    if (m.normalized.startsWith("/c") && !TEMPLATE_DEFINITIONS[base]) {
    const medRow = processMedicationCommand(m.text);
    if (medRow) TABLE_ROW_BUFFER.push(medRow);
    }
    });

    // 3. FINAL FLUSH
    insertDiagnosesIntoDocument();
    insertTemplatesIntoDocument(); 

    // 4. GLOBAL CLEANUP PASS
    if (cleanupQueue.length > 0) {
    const uniqueKeys = [...new Set(cleanupQueue)];
    uniqueKeys.forEach(key => {
    clearSectionByHeaderKey(key);
    });
    }

    if (TABLE_ROW_BUFFER.length > 0) generateMedicineTableFromBuffer();
    }

  // Expand Keywords
    function expandKeywords() {
    const body = DocumentApp.getActiveDocument().getBody();
    const combinedPattern = '\\/[a-zA-Z0-9]+(\\[.*?\\])*'; 
    let searchResult = null;
    const matches = [];

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

    for (let i = matches.length - 1; i >= 0; i--) {
    const m = matches[i];
    const parent = m.element.getParent();
    m.element.deleteText(m.start, m.end);
    if (parent.getType() === DocumentApp.ElementType.PARAGRAPH && parent.asParagraph().getText().trim() === "") {
    try { if (body.getNumChildren() > 1) parent.removeFromParent(); } catch (e) {}
    }
    }

    if (matches.length > 0) runExpansionEngine(matches);
    }

    function expandKeywordsFromSidebar(keyword) {
    const normalized = keyword.toLowerCase().trim();
    if (normalized === "/generatemedicinetable") {
    insertDiagnosesIntoDocument();
    insertTemplatesIntoDocument();
    if (TABLE_ROW_BUFFER.length > 0) generateMedicineTableFromBuffer();
    return;
    }
    runExpansionEngine([{ text: keyword, normalized: normalized }]);
    }

    function escapeForRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    function insertTemplateAtIndex(body, template, insertIndex) {
    const paragraphs = template.text.split('\n');
    const insertedParagraphs = [];
    const g = getGrammar('wellness', template.plurality || 'singular', template.sex || 'male');

    const resolve = (keys) => (keys || []).map(key => {
    const entry = FORMAT_REGISTRY[key];
    if (!entry) return null;
    return typeof entry === 'function' ? entry(g || { he:'he', him:'him', his:'his', dog:'dog', has:'has' }) : entry;
    }).filter(Boolean);

    const resolvedBoldOnly = resolve(template.boldKeys);
    const resolvedBoldUnderline = resolve(template.boldUnderlineKeys);
    const resolvedItalic = resolve(template.italicKeys);
    const resolvedGreen = resolve(template.greenKeys);
    const resolvedRed = resolve(template.redKeys);
    const resolvedLinks = resolve(template.linkKeys);
    const resolvedTitle = resolve(template.titleKeys);
    const resolvedDoubleSpaced = resolve(template.doubleSpacedKeys);

    paragraphs.forEach((paraText, i) => {
    const p = body.insertParagraph(insertIndex + i, paraText);
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY).setIndentFirstLine(36);
    let spacing = template.blockLineSpacing || 2.0;
    if (resolvedTitle.some(t => paraText.includes(t))) {
    p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    spacing = STYLE_REGISTRY.title.lineSpacing;
    }
    if (resolvedDoubleSpaced.some(ds => paraText.includes(ds))) {
    spacing = STYLE_REGISTRY.doubleSpaced.lineSpacing;
    }
    p.setLineSpacing(spacing);
    insertedParagraphs.push(p);
    });

    insertedParagraphs.forEach(p => {
    const t = p.editAsText();
    const text = p.getText();
    if (!text.trim()) return;
    resolvedTitle.forEach(titleText => { if (text.includes(titleText)) t.setFontSize(0, text.length - 1, 20); });
    resolvedBoldOnly.forEach(str => applyFormattingOnce(t, str, STYLE_REGISTRY.bold));
    resolvedBoldUnderline.forEach(str => applyFormattingOnce(t, str, STYLE_REGISTRY.boldUnderline));
    resolvedItalic.forEach(str => applyFormattingOnce(t, str, STYLE_REGISTRY.italic));
    resolvedGreen.forEach(str => applyFormattingOnce(t, str, STYLE_REGISTRY.green));
    resolvedRed.forEach(str => applyFormattingOnce(t, str, STYLE_REGISTRY.red));
    resolvedLinks.forEach(link => {
    let start = text.indexOf(link.text);
    while (start !== -1) {
    t.setLinkUrl(start, start + link.text.length - 1, link.url);
    start = text.indexOf(link.text, start + link.text.length);
    }
    });
    });

    if (template.table) {
    template.table.data.forEach((row, idx) => {
    TABLE_ROW_BUFFER.push({ rowData: row, color: template.table.colorRows ? template.table.colorRows[idx] : null });
    });
    }
    return insertedParagraphs;
    }

  // Search & Destroy
    function clearSectionByHeaderKey(headerKey) {
    const body = DocumentApp.getActiveDocument().getBody();
    const paragraphs = body.getParagraphs();
    const rawEntry = FORMAT_REGISTRY[headerKey];
    if (!rawEntry) return;

    // Resolve target text if it's a function
    const targetHeader = (typeof rawEntry === 'function') 
    ? rawEntry({ dog: 'dog', has: 'has', him: 'him', his: 'his', resists: 'resists' }) 
    : rawEntry;

    function isHeader(para) {
    const textObj = para.editAsText();
    const text = para.getText();
    if (!text.length) return false;
    for (let i = 0; i < text.length; i++) {
    if (textObj.isBold(i)) return true;
    }
    return false;
    }

    function isNextSubjectHeader(para, currentTarget) {
      const text = para.getText().trim();
      // Only treat as the "next" header if it has a colon AND isn't the one we just found
      if (!isHeader(para) || !text.includes(":") || text.startsWith(currentTarget)) return false;
      return true;
    }

    let startIdx = -1;
    let endIdx = -1;

    for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const rawText = p.getText();
    const text = rawText.replace(/\u00A0/g, ' ').trim(); 

    // Find start (using your working logic)
    if (startIdx === -1) {
    if (text.startsWith(targetHeader) && isHeader(p)) {
    startIdx = body.getChildIndex(p);
    }
    } 
    // Find end
    else {
      if (isNextSubjectHeader(p, targetHeader)) {
        endIdx = body.getChildIndex(p);
        break;
      }
    }
    }

    if (startIdx !== -1 && endIdx === -1) endIdx = body.getNumChildren();

    if (startIdx !== -1 && endIdx > startIdx) {
    if (endIdx - startIdx > 50) return; // Safety
    for (let j = endIdx - 1; j >= startIdx; j--) {
    try { body.removeChild(body.getChild(j)); } catch (e) {}
    }
    }
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
