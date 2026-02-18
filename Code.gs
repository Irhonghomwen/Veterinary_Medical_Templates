function onOpen() {
  DocumentApp.getUi()
    .createMenu('Medical Templates')
    .addItem('Open Template Picker')
    .addItem('Expand Keywords', 'expandKeywords')
    .addToUi();
}

/* ------------------ SHARED WELLNESS HELPERS ------------------ */

// Shared dental product links
const DENTAL_LINKS_SMALL = [
  { text: 'small dog toothbrush', url: 'https://www.chewy.com/virbac-cet-dog-cat-toothbrush/dp/42659' },
  { text: 'animal safe toothpaste', url: 'https://www.chewy.com/virbac-cet-enzymatic-poultry-flavor/dp/41660' }
];

const DENTAL_LINKS_LARGE = [
  { text: 'finger toothbrush', url: 'https://www.amazon.com/Virbac-C-E-T-Oral-Hygiene-Kit/dp/B004ULZ2PI' },
  { text: 'canine toothbrush', url: 'https://www.chewy.com/virbac-cet-dual-ended-dog-cat/dp/42658' },
  { text: 'animal safe toothpaste', url: 'https://www.chewy.com/virbac-cet-enzymatic-poultry-flavor/dp/41660' }
];

// Shared 8-week wellness medication table
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

/* ------------------ SHARED BOLD / UNDERLINE HELPERS ------------------ */

const BOLD_UNDERLINE_8WK_COMMON = [
  'Because your dog is still getting vaccines to better his immune system, it’s best to keep him away from dog parks & other dogs that aren’t part of your household until two weeks after he has finished his puppy vaccines.',
  'Keep him away from dog parks, training facilities, and other dogs until then.',
  'Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads).',
  'immediately',
  'These reactions are rare & not expected to occur in your dog.',
  'Because your dog is still growing, you will need to come back once a month to have him weighed & get the appropriate dose of preventative.',
  'It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites.',
  'The best way to keep your dog’s teeth healthy is to brush them daily for 10 seconds total',
  '(make sure xylitol isn’t listed as an ingredient),',
  'Brushing the outside for 1.5 seconds is more than enough.'
];

const BOLD_UNDERLINE_NEUTER_SMALL = [
  'It is recommended you have him neutered once he is 6 months old if you don’t intend to breed him.'
];

const BOLD_UNDERLINE_NEUTER_LARGE = [
  'It is recommended you have him neutered once he is 10 - 12 months old if you don’t intend to breed him.'
];

// Convenience builders
function getBoldUnderline8WkSmallMale() {
  return [...BOLD_UNDERLINE_8WK_COMMON, ...BOLD_UNDERLINE_NEUTER_SMALL];
}

function getBoldUnderline8WkLargeMale() {
  return [...BOLD_UNDERLINE_8WK_COMMON, ...BOLD_UNDERLINE_NEUTER_LARGE];
}

function getBoldUnderline8WkFemale() {
  return BOLD_UNDERLINE_8WK_COMMON;
}

/* ------------------ 8-WEEK WELLNESS TEMPLATE GENERATOR ------------------ */
function generate8WkWellnessTemplate(size, sex) {
  const pronoun = sex === 'female'
    ? { he: 'she', him: 'her', his: 'her' }
    : { he: 'he', him: 'him', his: 'his' };

  // Neuter paragraph for small vs large breed
  const spayorneuterText = sex === 'female'
   ? `Spay: It is recommended you have your dog spayed at 6 months if you do not intend to breed her. While it isn’t wrong to keep her intact, intact females are at risk of developing several life threatening diseases such as breast cancer, diabetes, & pyometra. 1 in 4 female dogs will get breast cancer if they are not spayed by their second heat cycle. In dogs there is a 50% chance breast cancer spreads throughout the body. The earliest time to have your dog spayed is when she is 6 months old before her first heat cycle. This will give her body enough time to grow while also reducing the risk of diseases that are associated with spaying too early. Even if she has already had her second heat cycle, spaying is still recommended since many mammary tumors are stimulated by estrogen & pyometra is still a possibility.`
  : size === 'large'
  ? `Neuter: On physical exam I was able to identify both of your dog’s testicles in ${pronoun.his} scrotum. It is recommended you have ${pronoun.him} neutered once ${pronoun.he} is 10 - 12 months old if you don’t intend to breed ${pronoun.him}. By 10 - 12 months of age, large breed dogs such as ${pronoun.him} have already received all the testosterone they need in order to grow normally. While it isn’t wrong to keep ${pronoun.him} intact, neutering ${pronoun.him} reduces or completely eradicates the risk of certain diseases such as prostatitis, several types of cancer, & hernias to name a few.`
  : `Neuter: On physical exam I was able to identify both of your dog’s testicles in ${pronoun.his} scrotum. It is recommended you have ${pronoun.him} neutered once ${pronoun.he} is 6 months old if you don’t intend to breed ${pronoun.him}. By 6 months of age, smaller breed dogs such as ${pronoun.him} have already received all the testosterone they need in order to grow normally. While it isn’t wrong to keep ${pronoun.him} intact, neutering ${pronoun.him} reduces or completely eradicates the risk of certain diseases such as prostatitis, several types of cancer, & hernias to name a few.`;

// Determine dental products based on size and sex
let dentalProducts = [];

if (sex === 'female') {
  dentalProducts = ['small dog toothbrush', 'finger toothbrush', 'canine toothbrush', 'animal safe toothpaste'];
} else if (size === 'large') {
  dentalProducts = ['finger toothbrush', 'canine toothbrush', 'animal safe toothpaste'];
} else {
  // small male
  dentalProducts = ['small dog toothbrush', 'animal safe toothpaste'];
}

  // Exact text for small male 8-week wellness, only substituting pronouns
  const text = `Vaccines: Your dog has received ${pronoun.his} first round of puppy vaccines today. Because of the antibodies that ${pronoun.he} received from ${pronoun.his} mother’s milk, the vaccines won’t provide full immunity until 16 weeks of age when most of the mother’s antibodies have disappeared. For that reason, it’s important to booster them every 3 - 4 weeks as your dog’s immune system slowly takes over.
Because your dog is still getting vaccines to better ${pronoun.his} immune system, it’s best to keep ${pronoun.him} away from dog parks & other dogs that aren’t part of your household until two weeks after ${pronoun.he} has finished ${pronoun.his} puppy vaccines.
The initial distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine was given as a combo shot in the left hindlimb. The 1 year bordetella vaccine was given orally. You may notice that after vaccination your dog is more tired than usual, eats less, or is sore at the injection site, & this is perfectly normal.
Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads). If you ever notice any of these within 24 hours of vaccination, bring your dog back immediately for treatment during normal business hours or your nearest emergency animal hospital. These reactions are rare & not expected to occur in your dog.
Heartworms: Heartworms are spread by mosquitoes which don’t die in the Texas "winter", so our pets are at risk of infection year round. Furthermore, heartworms can be fatal & there is a risk of death even with proper treatment. Prevention is easier, cheaper, & less stressful than treatment, so it is recommended you keep your dog on monthly preventatives such as Heartgard, Simparica Trio, Revolution, etc. Depending on the brand, they can protect your dog from heartworms, fleas, ticks, & common intestinal parasites with a single treatment. These can be given orally or topically & are generally well tolerated. Because your dog is still growing, you will need to come back once a month to have ${pronoun.him} weighed & get the appropriate dose of preventative.
${spayorneuterText}
Food: A high quality diet is the best way to keep your dog healthy. Any puppy diet from Hill’s Science Diet, Purina Pro Plan, or Royal Canin are all acceptable. A puppy diet is advised until your dog is a year old at which point you can transition to an adult diet. Dry food & wet food are both appropriate to feed. It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites. Follow the instructions on the back of the bag/can for a dog of ${pronoun.his} weight.
Dental care: The best way to keep your dog’s teeth healthy is to brush them daily for 10 seconds total using a ${dentalProducts.slice(0, -1).join(', ')} & ${dentalProducts[dentalProducts.length - 1]}. Animal safe toothpaste such as C.E.T. can be purchased from the clinic or from online stores. Getting your dog used to having ${pronoun.his} teeth brushed early will improve ${pronoun.his} overall health.
You can start by having ${pronoun.him} eat peanut butter (make sure xylitol isn’t listed as an ingredient), wet food, or treats off the toothbrush every day for a week, then applying the pet safe toothpaste & letting ${pronoun.him} lick it off every day for a week. Finally, gently brush ${pronoun.his} teeth with the toothpaste. Brushing the outside for 1.5 seconds is more than enough.
If your dog resists having ${pronoun.his} teeth brushed, dental cleanings can be performed under general anesthesia every few years as necessary for ${pronoun.his} teeth. Dental chews and water additives can also help slow down dental accumulation. You can find a list of products that have proven efficacy on the Veterinary Oral Health Council website.
Next appointment: Bring your dog back in 3 - 4 weeks for ${pronoun.his} next round of puppy vaccines.`;

  return {
    text,
    boldOnly: [
      'Vaccines:',
      'The initial distemper, adenovirus, parvovirus, & parainfluenza (DAPP) vaccine',
      'The 1 year bordetella vaccine',
      'Heartworms:',
      'Neuter:',
      'Spay',
      'Food:',
      'Dental care:',
      'Next appointment:'
    ],
    boldUnderline: [
      // Ensure your requested line is bold + underlined
      'Because your dog is still getting vaccines to better ' + pronoun.his + ' immune system, it’s best to keep ' + pronoun.him + ' away from dog parks & other dogs that aren’t part of your household until two weeks after ' + pronoun.he + ' has finished ' + pronoun.his + ' puppy vaccines.',
      'Watch out for severe vaccine reactions including swelling/pain at the vaccine sites, vomiting, diarrhea, extreme lethargy, or fever (excessive panting/sweating from the paw pads).',
  'immediately',
  'These reactions are rare & not expected to occur in your dog.',
  'Because your dog is still growing, you will need to come back once a month to have ' + pronoun.him + ' weighed & get the appropriate dose of preventative.',
  'It is recommended you have ' + pronoun.him + ' neutered once ' + pronoun.he + ' is ' + (size === 'large' ? '10 - 12 months' : '6 months') + ' old if you don’t intend to breed ' + pronoun.him + '.',
  'It is recommended you have your dog spayed at 6 months if you do not intend to breed her.',
  '1 in 4 female dogs will get breast cancer if they are not spayed by their second heat cycle. In dogs there is a 50% chance breast cancer spreads throughout the body.',
  'Even if she has already had her second heat cycle, spaying is still recommended since many mammary tumors are stimulated by estrogen & pyometra is still a possibility.',
  'It is not recommended to feed grain free or raw diets due to the increased risk of disease and parasites.',
  'The best way to keep your dog’s teeth healthy is to brush them daily for 10 seconds total using a small dog toothbrush & animal safe toothpaste.',
  'The best way to keep your dog’s teeth healthy is to brush them daily for 10 seconds total using a finger toothbrush, canine toothbrush & animal safe toothpaste.',
  'The best way to keep your dog’s teeth healthy is to brush them daily for 10 seconds total using a small dog toothbrush, finger toothbrush, canine toothbrush & animal safe toothpaste.',
  '(make sure xylitol isn’t listed as an ingredient),',
  'Brushing the outside for 1.5 seconds is more than enough.'
    ],
    greenText: [],
    links: [
      { text: 'Hill’s Science Diet', url: 'https://www.hillspet.com/dog-food#puppy%7Csd' },
      { text: 'Purina Pro Plan', url: 'https://www.purina.com/pro-plan/products/puppy-food' },
      { text: 'Royal Canin', url: 'https://www.royalcanin.com/us/dogs/products/retail-products?lifestage=baby%7Cpuppy' },
      { text: 'small dog toothbrush', url: 'https://www.chewy.com/virbac-cet-dog-cat-toothbrush/dp/42659' },
      { text: 'finger toothbrush', url: 'https://www.amazon.com/Virbac-C-E-T-Oral-Hygiene-Kit/dp/B004ULZ2PI' },
      { text: 'canine toothbrush', url: 'https://www.chewy.com/virbac-cet-dual-ended-dog-cat/dp/42658'},
      { text: 'animal safe toothpaste', url: 'https://www.chewy.com/virbac-cet-enzymatic-poultry-flavor/dp/41660' }
    ],
    table: get8WeekWellnessTable()
  };
}

/* ------------------ TEMPLATE DEFINITIONS ------------------ */

const TEMPLATES_LEGACY_CCL = {
  text: `Cranial cruciate ligament tear: Your dog has a torn cranial cruciate ligament, often called an ACL (humans) or CCL (animals). This ligament is found in the knee and, unlike humans, the most common cause is often genetic and due to how the knee is shaped in dogs as opposed to rough play or exercise. Again, this is NOT caused by inappropriate animal care, excessive exercise, or anything that was or wasn’t done at home.
Symptoms include limping, toe touching lameness, and refusal to use the affected leg that persists longer than 1 - 2 weeks. Diagnosis is through the x-rays that were performed and show the torn ligament. Treatment is best done through surgery. The gold standard is the tibial plateau leveling osteotomy (TPLO) which currently provides the best repair available. Alternatively, extracapsular repair can be performed as a more affordable option for dogs under 20 lbs.
It is strongly advised that you go forward with surgical repair via the TPLO surgery. However, at this time you’ve elected to go forward with medical management. You will need to exercise restrict your dog and use pain medication for up to 6 weeks. The purpose of this is to allow scarring to occur over the joint. While this doesn’t completely fix the problem, it limits the pain & inflammation. 
Because CCL tear occurs due to genetics rather than acute trauma, dogs with a torn CCL tend to have problems in the other knee within 6 - 12 months of diagnosis. You can learn more from the Ruptured Cranial Cruciate Ligaments in Dogs article by Veterinary Partner. You can also learn more about surgical repairs from the Cranial Cruciate Ligament Tear: Tibial Plateau Leveling Osteotomy (TPLO) and Cranial Cruciate Ligament Repair: Extracapsular Repair articles by VCA.`,
  boldOnly: ['Cranial cruciate ligament tear:'],
  boldUnderline: [
    'common cause',
      'Again, this is NOT caused by inappropriate animal care, excessive exercise, or anything that was or wasn’t done at home.',
      'Symptoms',
      'Diagnosis',
      'Treatment',
      'It is strongly advised that you go forward with surgical repair via the TPLO surgery. However, at this time you’ve elected to go forward with medical management.',
      'dogs with a torn CCL tend to have problems in the other knee within 6 - 12 months of diagnosis.'
  ],
  greenText: ['common cause', 'Symptoms', 'Diagnosis', 'Treatment'],
  links: [
    { text: 'Ruptured Cranial Cruciate Ligaments in Dogs article', url: 'https://veterinarypartner.vin.com/default.aspx?pid=19239&id=4952244' },
    { text: 'Cranial Cruciate Ligament Tear: Tibial Plateau Leveling Osteotomy (TPLO)', url: 'https://vcahospitals.com/know-your-pet/cranial-cruciate-ligament-repair-tibial-plateau-leveling-osteotomy-tplo' },
    { text: 'Cranial Cruciate Ligament Repair: Extracapsular Repair', url: 'https://vcahospitals.com/cranial-cruciate-ligament-repair-extracapsular-repair-and-tightrope-procedure' }
  ],
  table: {
      rows: 3,
      cols: 4,
      firstLineBoldUnderlineInstructions: true,
      data: [
        ['Medication', 'Instructions', 'Class', 'Side Effects'],
        ['Rimadyl 25mg (carprofen)', 'Starting today\nGive your dog 1 tablet by mouth with food every 12 hours for treatment of pain and inflammation. Give to completion.', 'Non-steroidal anti-inflammatory drug', 'May cause vomiting and diarrhea (less common if given with meals)'],
        ['Meloxicam 7.5mg', 'Starting today\nGive your dog 1 tablet by mouth with food every 24 hours for treatment of pain and inflammation. Give to completion.', 'Non-steroidal anti-inflammatory drug', 'May cause vomiting and diarrhea (less common if given with meals)']
      ],
      boldTopRow: true,
      
    }
};

const TEMPLATE_DEFINITIONS = {
  '/cCCL1MM': () => ({
    ...TEMPLATES_LEGACY_CCL
  }),

  '/c8wksSmallMale': () => generate8WkWellnessTemplate('small', 'male'),
  '/c8wksLargeMale': () => generate8WkWellnessTemplate('large', 'male'),
  '/c8wksFemale': () => generate8WkWellnessTemplate('small', 'female'),
};

function resolveTemplate(keyword) {
  const factory = TEMPLATE_DEFINITIONS[keyword];
  if (!factory) {
    throw new Error('Template not found: ' + keyword);
  }
  return factory();
}

/* ------------------ EXPAND KEYWORDS ------------------ */
function expandKeywords() {
  const body = DocumentApp.getActiveDocument().getBody();

  let foundAny = false;
  for (const key in TEMPLATES) {
    const regex = '(?i)' + key;
    let searchResult = null;

    while (searchResult = body.findText(regex, searchResult)) {
      foundAny = true;
      const textElement = searchResult.getElement().asText();
      textElement.deleteText(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive());
      const insertIndex = body.getChildIndex(textElement.getParent());
      insertTemplateAtIndex(body, resolveTemplate(key), insertIndex);
    }
  }

  if (!foundAny) {
    DocumentApp.getUi().alert('No recognized keywords found.');
  }
}

/* ------------------ INSERT TEMPLATE AT INDEX ------------------ */
function insertTemplateAtIndex(body, template, insertIndex) {
  const paragraphs = template.text.split('\n');
  const insertedParagraphs = [];

  paragraphs.forEach((paraText, i) => {
    const p = body.insertParagraph(insertIndex + i, paraText);
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    p.setIndentFirstLine(36);
    p.setLineSpacing(2);
    insertedParagraphs.push(p);
  });

  insertedParagraphs.forEach(p => {
    const t = p.editAsText();
    template.boldOnly.forEach(text => applyFormattingOnce(t, text, { bold: true }));
    template.boldUnderline.forEach(text => applyFormattingOnce(t, text, { bold: true, underline: true }));
    template.greenText.forEach(text => applyFormattingOnce(t, text, { color: '#6aa84f' }));
    template.links.forEach(link => {
  let startIndex = t.getText().indexOf(link.text);
  while (startIndex !== -1) {
    t.setLinkUrl(startIndex, startIndex + link.text.length - 1, link.url);
    // Look for next occurrence of the same text
    startIndex = t.getText().indexOf(link.text, startIndex + link.text.length);
  }
});
  });

  if (template.table) {
    insertTableAtIndex(template.table, insertIndex + insertedParagraphs.length);
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
function applyFormattingOnce(text, target, options) {
  const content = text.getText();
  const start = content.indexOf(target);
  if (start === -1) return;
  const end = start + target.length - 1;
  if (options.bold !== undefined) text.setBold(start, end, options.bold);
  if (options.underline !== undefined) text.setUnderline(start, end, options.underline);
  if (options.color) text.setForegroundColor(start, end, options.color);
}


