function onOpen() {
  const ui = DocumentApp.getUi();
  
  // Add menu items
  ui.createMenu("Dr. I.E. Osadiaye's Medical Templates")
    .addItem('Open Medical Template', 'showSidebar')
    .addItem('Expand Keywords', 'expandKeywords')
    .addToUi();

    showSidebar();
  }

/* ------------------ Shared Pronoun Helper ------------------ */
  function getPronoun(sex) {
  return sex === 'female'
  ? { he: 'she', He: 'She', him: 'her', Him: 'Her', his: 'her', His: 'Her' }
  : { he: 'he', He: 'He', him: 'him', Him: 'Him', his: 'his', His: 'His' };
  }

  const FORMAT_REGISTRY = {
/* ------------------ REGISTRY | Vaccines ------------------ */    
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

  IMMEDIATELY:
  'bring your dog back immediately for treatment',

  LEPTO_INITIAL:
  'the initial lepto vaccine',

  LEPTO_VXN:
  /the 1 year lepto vaccine/i,
  
  NEXT_APPOINTMENT_HEADER: 'Next appointment:',

  PUPPY_QUARANTINE: pronoun =>
  `Because your dog is still getting vaccines to better ${pronoun.his} immune system, it’s best to keep ${pronoun.him} away from dog parks & other dogs that aren’t part of your household until two weeks after ${pronoun.he} has finished ${pronoun.his} puppy vaccines.`,

  Quarantine_12WK: pronoun =>
  `Keep ${pronoun.him} away from dog parks, training facilities, and other dogs until then.`,

  Quarantine_16WK: pronoun =>
  `During this time, keep ${pronoun.him} away from dog parks & other dogs that aren’t part of your household.`,

  RABIES_1YR:
  'The 1 year rabies vaccine',

  RABIES_3YR:
  'The 3 year rabies vaccine',

  RARE_RXN:
  'These reactions are rare & not expected to occur in your dog.',

  VACCINES_HEADER:
  'Vaccines:',

  VXN_RXN:
  'Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads).',

/* ------------------ REGISTRY | Heartworms & Labwork ------------------ */
  HEARTWORMS_HEADER:
  'Heartworms:',

  HEARTWORM_TEST:
  'A heartworm test was performed on your dog. We will contact you in 3 - 4 business days with the results.',

  LAB_RESULTS:
  'Samples were drawn from your dog. You will receive a call in 3 - 4 business days with the results.',
  
  LABWORK:
  'Early detection labwork:',

/* ------------------ REGISTRY | Spay & Neuter ------------------ */
  LARGE_NEUTER: pronoun =>
  `It is recommended you have ${pronoun.him} neutered once ${pronoun.he} is 10 - 12 months old if you don’t intend to breed ${pronoun.him}.`,
  
  NEUTER_HEADER:
  'Neuter:',

  SMALL_NEUTER: pronoun =>
  `It is recommended you have ${pronoun.him} neutered once ${pronoun.he} is 6 months old if you don’t intend to breed ${pronoun.him}.`,
  
  SPAY1:
  'It is recommended you have your dog spayed at 6 months if you do not intend to breed her.',

  SPAY2:
  '1 in 4 female dogs will get breast cancer if they are not spayed by their second heat cycle. In dogs there is a 50% chance breast cancer spreads throughout the body.',

  SPAY3:
  'Even if she has already had her second heat cycle, spaying is still recommended since many mammary tumors are stimulated by estrogen & pyometra is still a possibility.',

  SPAY_HEADER: 'Spay:',

/* ------------------ REGISTRY | Diet ------------------ */
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

/* ------------------ REGISTRY | Dental ------------------ */
  DENTAL_BRUSHING:
  'Brushing the outside for 1.5 seconds is more than enough.',

  DENTAL_BRUSHING_CORE:
  /The best way to keep your dog’s teeth healthy is to brush them daily for 10 seconds total using a .*?toothpaste\./,
  
  DENTAL_HEADER: 'Dental care:',

  LARGE_TOOTHBRUSH_LINK: {
  text: 'medium/large dog toothbrush',
  url: 'https://a.co/d/0dgjdsJ8'
  },

  SMALL_DOG_TOOTHBRUSH_LINK: {
  text: 'small dog toothbrush',
  url: 'https://a.co/d/00v6EBUQ'
  },

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

/* ------------------ REGISTRY | General Illness ------------------ */
  COMMON_CAUSES:
  /Common causes/i,

  DIAGNOSIS:
  /Diagnosis/i,

  SYMPTOMS:
  /Symptoms/i,

  TREATMENT:
  /Treatment/i,

/* ------------------ REGISTRY | Weight Management ------------------ */
  DIET_LIFESPAN: pronoun =>
  `Helping ${pronoun.him} to lose weight can increase ${pronoun.his} life span by as much as 1 ½ years.`,

  DIET2_LIFESPAN: pronoun =>
  `Continuing to help ${pronoun.him} lose weight can extend ${pronoun.his} life span by as much as 1 ½ years.`,
  
  DIET_WEEKLY_GOAL: pronoun =>
  `We’re aiming to have ${pronoun.him} lose 1 - 2% of ${pronoun.his} body weight per week. If ${pronoun.he} begins losing more than that per week, increase the amount of food ${pronoun.he} gets.`,

  HEALTHY_WEIGHT_LIFESPAN: pronoun =>
  `Your dog is a healthy weight for a dog of ${pronoun.his} size.`,

  HILLS_DOG_DIET_DRY_LINK: {
  text: "Hill’s weight loss dry food",
  url: 'https://www.hillspet.com/dog-food?productform=dry&condition=weightmanagement'
  },

  HILLS_DOG_DIET_WET_LINK: {
  text: "Hill’s weight loss wet food",
  url: 'https://www.hillspet.com/dog-food?&productform=canned&productform=stew&condition=weightmanagement'
  },

  OVERWEIGHT_WARNING: pronoun =>
  `Your dog weighs more than the average dog of ${pronoun.his} size.`,

  OVERWEIGHT_DOG_SIGNS:
  `Signs of an overweight dog`,

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
  
  UNDERWEIGHT_DOG_SIGNS:
  `Signs of an underweight dog`,

  UNDERWEIGHT_HEADER:
  `Underweight:`,

  UNDERWEIGHT_LIFESPAN: pronoun =>
  `Helping ${pronoun.him} gain weight can increase ${pronoun.his} quality of life.`,

  UNDERWEIGHT_GOAL: pronoun =>
  `We're aiming to have ${pronoun.him} gain approximately 10% of ${pronoun.his} current weight. Failure to gain weight is concerning for disease and would prompt us to perform tests such as labwork, ultrasound, or x-rays.`,

  UNDERWEIGHT_WARNING: pronoun =>
  `Your dog weighs less than the average dog of ${pronoun.his} size.`,

  WEIGHT_HEADER:
  'Weight:',

/* ------------------ REGISTRY | Immunology ------------------ */
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
  
  LEPTO_WARNING:
  `Lepto can quickly kill our pets & can be spread to humans.`,

  PARVO_LINK: {
  text: `parvovirus article`,
  url: `https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4951460`,
  },

  RABIES_CURE:
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

  };

/* ------------------ Heartworm Prevention Table ------------------ */
  function get8WeekWellnessTable() {
  return {
    rows: 6,
    cols: 4,
    data: [
      ['Medication', 'Instructions', 'Class', 'Side Effects'],
      [
        'Heartgard',
        'Starting tomorrow\nGive your dog 1 chewable tablet every 30 days for prevention of heartworms and common intestinal parasites. ',
        'Parasiticide',
        'Rarely causes vomiting or diarrhea'
      ],
      [
        'Nexgard',
        'Starting tomorrow\nGive your dog 1 chewable tablet every 30 days for prevention of fleas and ticks. ',
        'Parasiticide',
        'Rarely causes vomiting or diarrhea'
      ],
      [
        'Proheart 12 (moxidectin)',
        'Given in clinic\nMedicine injected beneath your dog’s skin to prevent heartworm infection for 12 months. ',
        'Antiparasitic',
        'May cause lethargy, decreased appetite, or vomiting'
      ],
      [
        'Revolution',
        'Starting today\nApply contents between your dog’s ears every 30 days for prevention of heartworms, fleas, ticks, and common intestinal parasites.',
        'Parasiticide',
        'Rarely causes redness of the skin or hair loss'
      ],
      [
        'Simparica Trio',
        'Starting tomorrow\nGive your dog 1 chewable tablet by mouth every 30 days for prevention of heartworms, fleas, ticks, & common intestinal parasites.',
        'Antiparasitic',
        'Rarely causes vomiting, diarrhea, or neurologic abnormalities'
      ]
    ],
    boldTopRow: true,
    colorRows: { 3: '#a4c2f4' },
    firstLineBoldUnderlineInstructions: true
  };
  }

/* ------------------ 8 Week Wellness Template ------------------ */
  function generate8WkWellnessTemplate(size, sex) {
  const pronoun = getPronoun(sex);

  // Neuter paragraph for small vs large breed
  const spayorneuterText = sex === 'female'
   ? `Spay: It is recommended you have your dog spayed at 6 months if you do not intend to breed her. While it isn’t wrong to keep her intact, intact females are at risk of developing several life threatening diseases such as breast cancer, diabetes, & pyometra. 1 in 4 female dogs will get breast cancer if they are not spayed by their second heat cycle. In dogs there is a 50% chance breast cancer spreads throughout the body. The earliest time to have your dog spayed is when she is 6 months old before her first heat cycle. This will give her body enough time to grow while also reducing the risk of diseases that are associated with spaying too early. Even if she has already had her second heat cycle, spaying is still recommended since many mammary tumors are stimulated by estrogen & pyometra is still a possibility.`
  : size === 'large'
  ? `Neuter: On physical exam I was able to identify both of your dog’s testicles in ${pronoun.his} scrotum. It is recommended you have ${pronoun.him} neutered once ${pronoun.he} is 10 - 12 months old if you don’t intend to breed ${pronoun.him}. By 10 - 12 months of age, large breed dogs such as ${pronoun.him} have already received all the testosterone they need in order to grow normally. While it isn’t wrong to keep ${pronoun.him} intact, neutering ${pronoun.him} reduces or completely eradicates the risk of certain diseases such as prostatitis, several types of cancer, & hernias to name a few.`
  : `Neuter: On physical exam I was able to identify both of your dog’s testicles in ${pronoun.his} scrotum. It is recommended you have ${pronoun.him} neutered once ${pronoun.he} is 6 months old if you don’t intend to breed ${pronoun.him}. By 6 months of age, smaller breed dogs such as ${pronoun.him} have already received all the testosterone they need in order to grow normally. While it isn’t wrong to keep ${pronoun.him} intact, neutering ${pronoun.him} reduces or completely eradicates the risk of certain diseases such as prostatitis, several types of cancer, & hernias to name a few.`;

  // Determine dental products based on size and sex
  let dentalProducts = [
    'small dog toothbrush',
    'medium/large dog toothbrush',
    'animal safe toothpaste'
  ];

  if (sex === 'female') {
  dentalProducts = ['small dog toothbrush', 'medium/large dog toothbrush', 'animal safe toothpaste'];
  } else if (size === 'large') {
  dentalProducts = ['medium/large dog toothbrush', 'animal safe toothpaste'];
  } else {
  // small male
  dentalProducts = ['small dog toothbrush', 'animal safe toothpaste'];
  }

  // Exact text for small male 8-week wellness, only substituting pronouns
    const text = [
  `Vaccines: Your dog has received ${pronoun.his} first round of puppy vaccines today. Because of the antibodies that ${pronoun.he} received from ${pronoun.his} mother’s milk, the vaccines won’t provide full immunity until 16 weeks of age when most of the mother’s antibodies have disappeared. For that reason, it’s important to booster them every 3 - 4 weeks as your dog’s immune system slowly takes over.`,
  `Because your dog is still getting vaccines to better ${pronoun.his} immune system, it’s best to keep ${pronoun.him} away from dog parks & other dogs that aren’t part of your household until two weeks after ${pronoun.he} has finished ${pronoun.his} puppy vaccines.`,
  `The initial distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo shot in the left hindlimb. The 1 year bordetella vaccine was given orally. You may notice that after vaccination your dog is more tired than usual, eats less, or is sore at the injection site, & this is perfectly normal.`,
  `Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your dog back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your dog.`,
  `Heartworms: Heartworms are spread by mosquitoes which don’t die in the Texas "winter", so our pets are at risk of infection year round. Furthermore, heartworms can be fatal & there is a risk of death even with proper treatment. Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your dog on monthly preventatives such as Heartgard, Simparica Trio, Revolution, etc. Depending on the brand, they can protect your dog from heartworms, fleas, ticks, & common intestinal parasites with a single treatment. These can be given orally or topically & are generally well tolerated. Because your dog is still growing, you will need to come back once a month to have ${pronoun.him} weighed & get the appropriate dose of preventative.`,
  `${spayorneuterText}`,
  `Food: A high quality diet is the best way to keep your dog healthy. Any puppy diet from Hill’s Science Diet (Hill's puppy dry food or Hill's puppy wet food), Purina Pro Plan (Purina puppy dry food or Purina puppy wet food), or Royal Canin (RC puppy dry food or RC puppy wet food) are all acceptable. A puppy diet is advised until your dog is a year old at which point you can transition to an adult diet. Dry food & wet food are both appropriate to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag/can for a dog of ${pronoun.his} weight.`,
  `Dental care: The best way to keep your dog’s teeth healthy is to brush them daily for 10 seconds total using a ${dentalProducts.slice(0, -1).join(', ')} & ${dentalProducts[dentalProducts.length - 1]}. Animal safe toothpaste such as C.E.T. can be purchased from the clinic or from online stores. Getting your dog used to having ${pronoun.his} teeth brushed early will improve ${pronoun.his} overall health.`,
  `You can start by having ${pronoun.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then applying the pet safe toothpaste & letting ${pronoun.him} lick it off every day for a week. Finally, gently brush ${pronoun.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
  `If your dog resists having ${pronoun.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${pronoun.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`,
  
  `Next appointment: Bring your dog back in 3 - 4 weeks for ${pronoun.his} next round of puppy vaccines.`
  ].join('\n');

  return {
    sex,
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
      // Ensure your requested line is bold + underlined
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
    
    greenKeys: [

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

    table: get8WeekWellnessTable()
  };
  }

/* ------------------ 12 Week Wellness Template ------------------ */
  function generate12WkWellnessTemplate(size, sex) {
  const pronoun = getPronoun(sex);

  // 1. Start from the 8-week template
  const template = generate8WkWellnessTemplate(size, sex);

  // 2. Replace ONLY the vaccine section
  const newVaccineText12wks = [
  `Vaccines: Your dog has received ${pronoun.his} next round of puppy vaccines today. Because of the antibodies that ${pronoun.he} received from ${pronoun.his} mother’s milk, ${pronoun.he} won’t have full immunity until 2 weeks after ${pronoun.his} last round of puppy vaccines. Keep ${pronoun.him} away from dog parks, training facilities, and other dogs until then.`,
  `The distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo shot with the initial lepto vaccine in the left hindlimb. You may notice that after vaccination your dog is more tired than usual, eats less, or is sore at the injection site, & this is perfectly normal.`,
  `Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your dog back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your dog.`
  ].join('\n');

  // 3. Replace the first vaccine block in the text
  template.text = template.text.replace(
    /Vaccines:[\s\S]*?These reactions are rare & not expected to occur in your dog\./,
    newVaccineText12wks
  );

  // 4. Replace ONLY the next appointment sentence
  template.text = template.text.replace(
    /Next appointment:[\s\S]*?\./,
    'Next appointment: Bring your dog back in 4 weeks for ' + pronoun.his + ' final round of puppy vaccines.'
  );

  // 5. Add to 8 week bold-only
  template.boldKeys = [
  ...(template.boldKeys || []),
  'DAPP_BOOSTER',
  'LEPTO_INITIAL',
  ];


  // 6. Bold/underline
  template.boldUnderlineKeys = template.boldUnderlineKeys.concat([
  'Quarantine_12WK',
  ]);

  
  return template;
  }

/* ------------------ 16 Week Wellness Template ------------------ */
  function generate16WkSmallMaleTemplate(size, sex) {
  const pronoun = getPronoun(sex);

  // 1. Start from the 8-week template
  const template = generate8WkWellnessTemplate(size, sex);

  // 2. Replace ONLY the vaccine section
  const newVaccineText16wks = [
  `Vaccines: Your dog has received ${pronoun.his} final round of puppy vaccines today. The antibodies ${pronoun.he} received from ${pronoun.his} mother’s milk have mostly decreased at this point, so ${pronoun.his} own immune system should be in full effect. Starting from today, ${pronoun.he} can get ${pronoun.his} vaccines once a year.`,
  `Over the next two weeks your dog will build up immunity to the vaccines given. During this time, keep ${pronoun.him} away from dog parks & other dogs that aren’t part of your household. Afterwards ${pronoun.he} is safe to interact with other dogs.`,
  `The 1 year rabies vaccine was given in the right hindlimb. The 1 year distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo shot with the 1 year lepto vaccine in the left hindlimb. You may notice that after vaccination your dog is more tired than usual, eats less, or is sore at the injection site, & this is perfectly normal.`,
  `Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your dog back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your dog.`
  ].join('\n');

  // 3. Replace the first vaccine block in the text
  template.text = template.text.replace(
    /^Vaccines:[\s\S]*?These reactions are rare & not expected to occur in your dog\./,
    newVaccineText16wks
  );

  // 4. Replace ONLY the next appointment line with gender-specific text
  template.text = template.text.replace(
  /Next appointment:[\s\S]*$/,
  `Next appointment: Bring your dog back once ${pronoun.he} is at least 6 months old for ${pronoun.his} ${sex === 'female' ? 'spay' : 'neuter'} & heartworm test. If you would not like to ${sex === 'female' ? 'spay' : 'neuter'} ${pronoun.him}, bring ${pronoun.him} back in 1 year for ${pronoun.his} annual adult vaccines.`
  );

  // 5. Add to 8 week bold-only
  template.boldKeys = [
  ...(template.boldKeys || []),
  'RABIES_1YR',
  'DAPP_1YR',
  'LEPTO_VXN',
  ];

  // 6. Bold/underline
  template.boldUnderlineKeys = template.boldUnderlineKeys.concat([
  'Quarantine_16WK',
  ]);

  return template;
  }

/* ------------------ Initial Adult Vaccine Template  ------------------ */
  function generateInitialAdultTemplate(sex) {
  const pronoun = getPronoun(sex);

  // Base template from 8-week wellness
  const template = generate8WkWellnessTemplate('small', sex);

  const dentalProducts = [
  'small dog toothbrush',
  'medium/large dog toothbrush',
  'animal safe toothpaste'
  ];

  // Adult vaccine text
  const adultVaccineText = [
  `Vaccines: Your dog has received ${pronoun.his} first round of adult vaccinations. The 1 year rabies vaccine was given in the right hindlimb. The initial distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo shot with the initial lepto vaccine in the left hindlimb. The 1 year bordetella vaccine was given orally. Your dog will need a booster of the DAPP and lepto vaccines in 3 - 4 weeks.`,
  `You may notice that after vaccination your dog is more tired than usual, eats less, or is sore at the injection site, & this is perfectly normal. Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your dog back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your dog.`,
  `Heartworms: A heartworm test was performed on your dog. We will contact you in 3 - 4 business days with the results. Heartworms are spread by mosquitoes which don’t die in the Texas "winter", so our pets are at risk of infection year round. Furthermore, heartworms can be fatal & there is a risk of death even with proper treatment. Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your dog on monthly preventatives such as Heartgard, Nexgard, Simparica Trio, Revolution, etc.`,
  `Early detection labwork: Yearly blood work is recommended for dogs the same as it is in humans and starts at 3 years of age. This lets us get a baseline for your pet and allows us to catch abnormalities before they’re noticeable outwardly. Depending on the panel run, this can check for issues in the liver, kidneys, thyroid, bladder, glucose, and many other organs and values. At 6 years of age, a larger panel for “senior” pets is advised.`,
  `Food: A high quality diet is the best way to keep your dog healthy. Food from Hill’s Science Diet (Hill's dog dry food or Hill's dog wet food), Purina Pro Plan (Purina dog dry food or Purina dog wet food), or Royal Canin (RC dog dry food or RC dog wet food) are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in dogs, so either are wonderful to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag/can for a dog of your dog’s weight.`,
  `Dental care: The best way to keep your dog’s teeth healthy is to brush them daily for 10 seconds total using a ${dentalProducts.slice(0, -1).join(', ')} & ${dentalProducts[dentalProducts.length - 1]}. Animal safe toothpaste such as C.E.T. can be purchased from the clinic or from online stores. Getting your dog used to having ${pronoun.his} teeth brushed early will improve ${pronoun.his} overall health.`,
  `You can start by having ${pronoun.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then applying the pet safe toothpaste & letting ${pronoun.him} lick it off every day for a week. Finally, gently brush ${pronoun.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
  `If your dog resists having ${pronoun.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${pronoun.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`,
  `Next appointment: Bring your dog back in 3 - 4 weeks for a booster of ${pronoun.his} vaccines.`
  ].join('\n');

  template.text = adultVaccineText;

  // Update bold formatting
  template.boldKeys = [
  ...(template.boldKeys || []),
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

  template.boldUnderlineKeys = [
    ...(template.boldUnderlineKeys || []),
      'HEARTWORM_TEST',
    ],

  // Update links from 8-week
  template.linkKeys = [
  ...(template.linkKeys || []),
  'HILLS_DOG_DRY_LINK',
  'HILLS_DOG_WET_LINK',
  'PURINA_DOG_DRY_LINK',
  'PURINA_DOG_WET_LINK',
  'ROYAL_CANIN_DOG_DRY_LINK',
  'ROYAL_CANIN_DOG_WET_LINK',
  ],

  // Table same as 8-week
  template.table = get8WeekWellnessTable();

  return template;
  }

/* ------------------ 1-Year Adult Vaccine Template ------------------ */
  function generate1YearAdultTemplate(sex) {
  const pronoun = getPronoun(sex);
    
  const dentalProducts = [
  'small dog toothbrush',
  'medium/large dog toothbrush',
  'animal safe toothpaste'
  ];

  const text = [ `Vaccines: Your dog has received ${pronoun.his} first round of adult vaccinations. Because you have kept to ${pronoun.his} vaccination schedule, ${pronoun.his} immune system will not need another booster until next year.`,
  `The 1 year rabies vaccine was given in the right hindlimb. The 1 year distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo shot with the 1 year lepto vaccine in the left hindlimb. The 1 year bordetella vaccine was given orally. You may notice that after vaccination your dog is more tired than usual, eats less, or is sore at the injection site, & this is perfectly normal.`,
  `Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your dog back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your dog.`,
  `Heartworms: A heartworm test was performed on your dog. We will contact you in 3 - 4 business days with the results. Heartworms are spread by mosquitoes which don’t die in the Texas "winter", so our pets are at risk of infection year round. Furthermore, heartworms can be fatal & there is a risk of death even with proper treatment. Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your dog on monthly preventatives such as Heartgard, Nexgard, Simparica Trio, Revolution, etc.`,
  `Early detection labwork: Samples were drawn from your dog. You will receive a call in 3 - 4 business days with the results. Yearly blood work is recommended for dogs the same as it is in humans for the sake of monitoring for abnormalities that aren’t visible from the outside. Depending on the panel run, this can check for issues in the liver, kidneys, thyroid, bladder, glucose, and many other organs and values. If no abnormalities are found, the results can be used as a baseline so that your dog’s overall health is closely monitored.`,
  `Food: A high quality diet is the best way to keep your dog healthy. If you haven’t already, you can transition ${pronoun.him} from ${pronoun.his} puppy diet to ${pronoun.his} adult diet. Food from Hill’s Science Diet (Hill's dog dry food or Hill's dog wet food), Purina Pro Plan (Purina dog dry food or Purina dog wet food), or Royal Canin (RC dog dry food or RC dog wet food) are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in dogs, so either are wonderful to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag/can for a dog of ${pronoun.his} weight.`,
  `Dental care: The best way to keep your dog’s teeth healthy is to brush them daily for 10 seconds total using a ${dentalProducts.slice(0, -1).join(', ')} & ${dentalProducts[dentalProducts.length - 1]}. Animal safe toothpaste such as C.E.T. can be purchased from the clinic or from online stores. Getting your dog used to having ${pronoun.his} teeth brushed early will improve ${pronoun.his} overall health.`,
  `You can start by having ${pronoun.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then applying the pet safe toothpaste & letting ${pronoun.him} lick it off every day for a week. Finally, gently brush ${pronoun.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.`,
  `If your dog resists having ${pronoun.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${pronoun.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.`,
  `Next appointment: Bring your dog back one year from today for ${pronoun.his} next annual vaccines.`
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

    greenKeys: [],

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
    
    table: get8WeekWellnessTable()
  };
  }

/* ------------------ 2-Year Adult Vaccine Template ------------------ */
  function generate2YearAdultTemplate(sex) {
  const template = generate1YearAdultTemplate(sex);

  template.boldKeys = [
  ...(template.boldKeys || []),
  'RABIES_3YR',
  'DAPP_3YR',
  ];

  const pronoun = getPronoun(sex);

  const adultFoodRecommendation = "A high quality diet is the best way to keep your dog healthy. Food from Hill’s Science Diet, Purina Pro Plan, or Royal Canin are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in dogs, so either are wonderful to feed.";

  // 2. Replace the 1 year vaccines with 3 year vaccines instead.
  template.text = template.text.replace(
  /Your dog has received (his|her) first round of adult vaccinations\./,
  `Your dog has received ${pronoun.his} annual adult vaccinations.`
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


  // 3. Remove puppy food transition.
  template.text = template.text.replace(
    /If you haven’t already, you can transition (him|her) from (his|her) puppy diet to (his|her) adult diet\.\s*/i,
    ''
  );

  return template;
  }

/* ------------------ 2-Year Lepto Vaccine Template ------------------ */

  function generate2YearLeptoTemplate(sex) {
  const template = generate2YearAdultTemplate(sex);
  const pronoun = getPronoun(sex);

  // Replace the vaccine listing paragraph ONLY
  template.text = template.text.replace(
    /The 3 year rabies vaccine was given in the right hindlimb\. The 3 year distemper, adenovirus, parvovirus, & parainfluenza \(DAPP\) vaccine was given as a combo shot with the 1 year lepto vaccine in the left hindlimb\. The 1 year bordetella vaccine was given orally\./,
    'The 1 year lepto vaccine was given in the left hindlimb. The 1 year bordetella vaccine was given orally.'
  );

  return template;
  }

/* ------------------ 7-Year Adult Vaccine Template ------------------ */

  function generate7YearAdultTemplate(sex) {
  const template = generate2YearAdultTemplate(sex);
  const pronoun = getPronoun(sex);

  // 1. Adjust senior diet wording
  template.text = template.text.replace(
  /Food: A high quality diet is the best way to keep your dog healthy\.[\s\S]*?can for a dog of (his|her) weight\./,
  `Food: A high quality diet is the best way to keep your dog healthy. Dogs that are older than 7 years are advised to be on a senior diet. Food from Hill’s Science Diet (Hill's senior dog dry food or Hill's senior dog wet food), Purina Pro Plan (Purina senior dog dry food or Purina senior dog wet food), or Royal Canin (RC senior dog dry food or RC senior dog wet food) are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in dogs, so either are wonderful to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag/can for a dog of his weight.`
  );

  // 3. Add senior dog links
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

/* ------------------ 7-Year Lepto Vaccine Template ------------------ */

  function generate7YearLeptoTemplate(sex) {
  const template = generate2YearAdultTemplate(sex);
  const pronoun = getPronoun(sex);

  // 1. Replace the vaccine listing paragraph ONLY
  template.text = template.text.replace(
    /The 3 year rabies vaccine was given in the right hindlimb\. The 3 year distemper, adenovirus, parvovirus, & parainfluenza \(DAPP\) vaccine was given as a combo shot with the 1 year lepto vaccine in the left hindlimb\. The 1 year bordetella vaccine was given orally\./,
    'The 1 year lepto vaccine was given in the left hindlimb. The 1 year bordetella vaccine was given orally.'
  );

  // 2. Adjust senior diet wording
  template.text = template.text.replace(
  /Food: A high quality diet is the best way to keep your dog healthy\.[\s\S]*?can for a dog of (his|her) weight\./,
  `Food: A high quality diet is the best way to keep your dog healthy. Dogs that are older than 7 years are advised to be on a senior diet. Food from Hill’s Science Diet (Hill's senior dog dry food or Hill's senior dog wet food), Purina Pro Plan (Purina senior dog dry food or Purina senior dog wet food), or Royal Canin (RC senior dog dry food or RC senior dog wet food) are all wonderful diets as they’re formulated by veterinary scientists. There is no significant difference between wet or dry food in dogs, so either are wonderful to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag/can for a dog of his weight.`
  );

  // 3. Add senior dog links
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

/* ------------------ Canine Diet Program | 1st ------------------ */

  function generateDietTemplate(sex) {
  const p = getPronoun(sex);

  const text = [`Weight: Your dog weighs more than the average dog of ${p.his} size. Ideally we would be able to feel ${p.his} ribs but not see them. Helping ${p.him} to lose weight can increase ${p.his} life span by as much as 1 ½ years. The best way to lose weight is through diet.`,
  `You can use the diet ${p.he} is currently on or you can use a prescription weight loss food from Hill’s Prescription Diet (Hill’s weight loss dry food or Hill’s weight loss wet food), Purina Pro Plan (Purina weight loss dry food or Purina weight loss wet food), or Royal Canin (RC weight loss dry food or RC weight loss wet food). Regardless, begin by measuring how much your dog eats using a measuring cup. Make sure to feed twice daily on a schedule rather than leaving food down at all times. If ${p.he} steals food from siblings, you may need to feed separately. Finally, decrease ${p.his} food by 10 - 25%.`,
  `We’re aiming to have ${p.him} lose 1 - 2% of ${p.his} body weight per week. If ${p.he} begins losing more than that per week, increase the amount of food ${p.he} gets. Another way you can help ${p.him} lose weight is by converting ${p.his} treats into healthy alternatives such as slices of apples, carrots, ice cubes, cucumbers, or green beans.`
  ].join('\n');

  return {
    sex,
    text,
    boldKeys: [
      'WEIGHT_HEADER',
    ],

    boldUnderlineKeys: [
      'DIET_LIFESPAN',
      'DIET_WEEKLY_GOAL',
      'OVERWEIGHT_WARNING',
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

/* ------------------ Canine Diet Program | 2nd, Continue ------------------ */

  function generateDiet2Template(sex) {
  const p = getPronoun(sex);

    const text = [
  `Weight: Congrats on helping ${p.him} lose weight! Continuing to help ${p.him} lose weight can extend ${p.his} life span by as much as 1 ½ years.`,
  `As a reminder, food from Hill’s Prescription Diet (Hill’s weight loss dry food or Hill’s weight loss wet food), Purina Pro Plan (Purina weight loss dry food or Purina weight loss wet food), or Royal Canin (RC weight loss dry food or RC weight loss wet food) can be used as needed. Otherwise, continue measuring how much ${p.he} eats using a measuring cup and feeding on a twice daily schedule rather than leaving food down at all times. Separate ${p.him} from siblings at meal time if necessary.`,
  `We’re aiming to have ${p.him} lose 1 - 2% of ${p.his} body weight per week. If ${p.he} begins losing more than that per week, increase the amount of food ${p.he} gets. Another way you can help ${p.him} lose weight is by converting ${p.his} treats into healthy alternatives such as slices of apples, carrots, ice cubes, cucumbers, or green beans.`
  ].join('\n');

  return {
    sex,
    text,
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

/* ------------------ Canine Healthy Weight ------------------ */

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
  
/* ------------------ Canine Underweight ------------------ */
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

    boldKeys: [
      'UNDERWEIGHT_HEADER',
    ],

    boldUnderlineKeys: [
      'UNDERWEIGHT_LIFESPAN',
      `UNDERWEIGHT_GOAL`,
      'UNDERWEIGHT_WARNING',
    ],

    greenKeys: [],
    linkKeys: []
  };
  }

/* ------------------ Canine Vaccine Information ------------------ */
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

    boldKeys: [
      'RABIES_HEADER',
      'DAPP_HEADER',
      'LEPTO_HEADER',
      'BORDETELLA_HEADER',
    ],

    boldUnderlineKeys: [
      `DAPP_WARNING`,
      `LEPTO_WARNING`
    ],

    greenKeys: [],

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

/* ------------------ TEMPLATE DEFINITIONS ------------------ */

const TEMPLATES_LEGACY_CCL = {
  text: ``,
  boldKeys: [],
  boldUnderlineKeys: [],
  greenKeys: [],
  linkKeys: [],
};

const RAW_TEMPLATE_DEFINITIONS = {
  '/cCCL1MM': () => ({
    ...TEMPLATES_LEGACY_CCL
  }),

  '/ieolibrary': function () {
  return {
    text: buildKeywordLibrary(),
    boldKeys: [],
    boldUnderlineKeys: [],
    greenKeys: [],
    linkKeys: []
  };
},

/* ------------------ Puppy Wellness ------------------ */
  '/c8wkssmallmale': () => generate8WkWellnessTemplate('small', 'male'),
  '/c8wkslargemale': () => generate8WkWellnessTemplate('large', 'male'),
  '/c8wksfemale': () => generate8WkWellnessTemplate('small', 'female'),

  '/c12wkssmallmale': () => generate12WkWellnessTemplate('small', 'male'),
  '/c12wkslargemale': () => generate12WkWellnessTemplate('large', 'male'),
  '/c12wksfemale': () => generate12WkWellnessTemplate('small', 'female'),

  '/c16wkssmallmale': () => generate16WkSmallMaleTemplate('small', 'male'),
  '/c16wkslargemale': () => generate16WkSmallMaleTemplate('large', 'male'),
  '/c16wksfemale': () => generate16WkSmallMaleTemplate('small', 'female'),

/* ------------------ Canine Adult Wellness ------------------ */

  '/cinitialadultmale': () => generateInitialAdultTemplate('male'),
  '/cinitialadultfemale': () => generateInitialAdultTemplate('female'),
  '/c1yearmale': () => generate1YearAdultTemplate('male'),
  '/c1yearfemale': () => generate1YearAdultTemplate('female'),
  '/c2yearmale': () => generate2YearAdultTemplate('male'),
  '/c2yearfemale': () => generate2YearAdultTemplate('female'),
  '/c2yearleptomale': () => generate2YearLeptoTemplate('male'),
  '/c2yearleptofemale': () => generate2YearLeptoTemplate('female'),
  '/c7yearmale': () => generate7YearAdultTemplate('male'),
  '/c7yearfemale': () => generate7YearAdultTemplate('female'),
  '/c7yearleptomale': () => generate7YearLeptoTemplate('male'),
  '/c7yearleptofemale': () => generate7YearLeptoTemplate('female'),

  /* ------------------ Canine Weight Management ------------------ */
  '/cdietmale': () => generateDietTemplate('male'),
  '/cdietfemale': () => generateDietTemplate('female'),
  '/cdiet2male': () => generateDiet2Template('male'),
  '/cdiet2female': () => generateDiet2Template('female'),
  '/cHealthyWeightMale': () => generateDogHealthyWeightTemplate('male'),
  '/cHealthyWeightFemale': () => generateDogHealthyWeightTemplate('female'),
  '/cUnderweightMale': () => generateDogUnderweightTemplate('male'),
  '/cUnderweightFemale': () => generateDogUnderweightTemplate('female'),

  /* ------------------ Immunology ------------------ */
  '/cVaccineInformation': () => generateDogVaccineInformationTemplate(),
};

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
function expandKeywords() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  const keys = Object.keys(TEMPLATE_DEFINITIONS)
    .map(k => escapeForRegex(k));

  if (keys.length === 0) return;

  const combinedPattern = '(?i)(' + keys.join('|') + ')';

  let searchResult = null;
  const matches = [];

  // --- Collect all matches first ---
  while ((searchResult = body.findText(combinedPattern, searchResult))) {
    matches.push(searchResult);
  }

  // --- Expand from bottom to top ---
  for (let i = matches.length - 1; i >= 0; i--) {
    const result = matches[i];
    const textElement = result.getElement().asText();

    const matchedText = textElement.getText().substring(
      result.getStartOffset(),
      result.getEndOffsetInclusive() + 1
    );

    const normalized = matchedText.trim().toLowerCase();
    const templateFn = TEMPLATE_DEFINITIONS[normalized];

    if (!templateFn) continue;

    textElement.deleteText(
      result.getStartOffset(),
      result.getEndOffsetInclusive()
    );

    const insertIndex = body.getChildIndex(textElement.getParent());

    insertTemplateAtIndex(body, templateFn(), insertIndex);
  }
}

/* ------------------ REGEX HELPERS ------------------ */
function escapeForRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/* ------------------ INSERT TEMPLATE AT INDEX ------------------ */
function insertTemplateAtIndex(body, template, insertIndex) {
  const paragraphs = template.text.split('\n');
  const insertedParagraphs = [];

  const pronoun = template.sex
  ? getPronoun(template.sex)
  : null;

const resolvedBoldOnly = [
  ...(template.boldOnly || []),
  ...(template.boldKeys || [])
    .map(key => FORMAT_REGISTRY[key])
    .filter(Boolean),
];

const resolvedBoldUnderline = [
  ...(template.boldUnderline || []),
  ...(template.boldUnderlineKeys || [])
    .map(key => {
      const entry = FORMAT_REGISTRY[key];
      if (!entry) return null;
      return typeof entry === 'function'
        ? entry(pronoun || { he:'he', him:'him', his:'his' })
        : entry;
    })
    .filter(Boolean),
];

const resolvedGreen = [
  ...(template.green || []),
  ...(template.greenKeys || [])
    .map(key => {
      const entry = FORMAT_REGISTRY[key];
      if (!entry) return null;
      return typeof entry === 'function'
        ? entry(pronoun || { he:'he', him:'him', his:'his' })
        : entry;
    })
    .filter(Boolean),
];

const resolvedRed = [
  ...(template.red || []),
  ...(template.redKeys || [])
    .map(key => {
      const entry = FORMAT_REGISTRY[key];
      if (!entry) return null;
      return typeof entry === 'function'
        ? entry(pronoun || { he:'he', him:'him', his:'his' })
        : entry;
    })
    .filter(Boolean),
];

const resolvedLinks = [
  ...(template.links || []), // optional legacy fallback
  ...(template.linkKeys || [])
    .map(key => FORMAT_REGISTRY[key])
    .filter(Boolean),
];

  // STEP 1: Insert paragraphs
  paragraphs.forEach((paraText, i) => {
    const p = body.insertParagraph(insertIndex + i, paraText);
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    p.setIndentFirstLine(36);
    p.setLineSpacing(2);
    insertedParagraphs.push(p);
  });

  // STEP 2: Apply formatting
  insertedParagraphs.forEach(p => {
    const t = p.editAsText();

    // Bold only
    resolvedBoldOnly.forEach(text =>
    applyFormattingOnce(t, text, { bold: true })
    );

    // Bold + underline (standalone cases)
    resolvedBoldUnderline.forEach(text =>
    applyFormattingOnce(t, text, { bold: true, underline: true })
    );

    // Green keys = bold + underline + green
    resolvedGreen.forEach(text =>
    applyFormattingOnce(t, text, {
    bold: true,
    underline: true,
    color: '#6aa84f'
    })
    );

    // Red keys = bold + underline + red
    resolvedRed.forEach(text =>
    applyFormattingOnce(t, text, {
    bold: true,
    underline: true,
    color: '#ff0000'
    })
    );

    // Links
    resolvedLinks.forEach(link => {
    let startIndex = t.getText().indexOf(link.text);

    while (startIndex !== -1) {
      t.setLinkUrl(
        startIndex,
        startIndex + link.text.length - 1,
        link.url
    );

    startIndex = t
      .getText()
      .indexOf(link.text, startIndex + link.text.length);
  }
});


  });

  // STEP 3: Insert table (if present)
  if (template.table) {
    insertTableAtIndex(
      template.table,
      insertIndex + insertedParagraphs.length
    );
  }

  return insertedParagraphs;
}

/* ------------------ TABLE INSERTION ------------------ */
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

/* ------------------ FORMATTING HELPERS ------------------ */
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
  if (style.color) textElement.setForegroundColor(start, end, style.color);
}

/* ------------------ SHOW SIDEBAR ------------------ */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle("Dr. I.E. Osadiaye's Medical Templates");
  DocumentApp.getUi().showSidebar(html);
}

/* ------------------ EXPAND KEYWORDS FROM SIDEBAR ------------------ */
function expandKeywordsFromSidebar(keyword) {
  const body = DocumentApp.getActiveDocument().getBody();
  const insertIndex = body.getNumChildren();
  insertTemplateAtIndex(body, resolveTemplate(keyword), insertIndex);
}
