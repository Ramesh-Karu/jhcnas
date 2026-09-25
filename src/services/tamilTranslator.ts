/**
 * Comprehensive Tamil-to-English translation, transliteration & intelligent column identifier sanitizer.
 * Handles inventory, accounting, ledger, school, student, and administrative terms.
 */

// Comprehensive Tamil dictionary for inventory, accounting, educational and tabular records
const TAMIL_DICTIONARY: Record<string, string> = {
  // Inventory & Stock Specific
  'பிரிவும் இலக்கமும்': 'section_and_number',
  'பிரிவும் இலக்கம்': 'section_and_number',
  'பிரிவு இலக்கம்': 'section_and_number',
  'பிரிவு மற்றும் இலக்கம்': 'section_and_number',
  'பேரேட்டில் காட்டியவாறான மீதி': 'ledger_balance',
  'பேரேட்டில் காட்டியவாறான': 'ledger_balance',
  'பேரேட்டு மீதி': 'ledger_balance',
  'பேரேடு': 'ledger',
  'உண்மையான கையிருப்பு மீதி': 'actual_balance_on_hand',
  'உண்மையான கையிருப்பு': 'actual_balance_on_hand',
  'கையிருப்பு மீதி': 'stock_on_hand_balance',
  'கையிருப்பு': 'stock_on_hand',
  'உபரி': 'surplus',
  'பற்றாக்குறை': 'deficiency',
  'குறிப்பு': 'remarks',
  'குறிப்புகள்': 'remarks',
  'விபரம்': 'details',
  'விவரம்': 'details',
  'பொருட்கள்': 'articles',
  'பொருட்கள் விபரம்': 'articles_description',
  'பொருள்': 'item',
  'விளக்கம்': 'description',
  'பொறுப்பாளர்': 'responsible_person',
  'பொறுப்பானவர்': 'responsible_person',
  'பொறுப்பு': 'responsibility',
  'கையொப்பம்': 'signature',
  'ஒப்பம்': 'signature',
  'தொடர்பு எண்': 'contact_number',
  'தொலைபேசி எண்': 'phone_number',
  'தொலைபேசி': 'phone',
  'மகிழகம்': 'mahilakam',
  'மகிலகம்': 'mahilakam',
  'ஸ்மார்ட் வகுப்பறை': 'smart_classroom',
  'பக்க எண்': 'page_no',
  'பக்கம்': 'page',
  'விலை': 'price',
  'பெறுமதி': 'valuation',
  'மதிப்பீடு': 'valuation',
  'மதிப்பு': 'value',
  'தொகை': 'amount',
  'எண்ணிக்கை': 'quantity',
  'அளவு': 'quantity',
  'மொத்தம்': 'total',
  'வரிசை எண்': 'serial_no',
  'இலக்கம்': 'number',
  'எண்': 'number',
  'பிரிவு': 'division',
  'வகுப்பு': 'grade',
  'துறை': 'department',
  'திணைக்களம்': 'department',
  'அலுவலகம்': 'office',
  'நூலகம்': 'library',
  'ஆய்வுகூடம்': 'lab',
  'ஆய்வகம்': 'lab',
  'விளையாட்டு': 'sports',
  'கணினி': 'ict_computer',
  'இயற்பியல்': 'physics',
  'பௌதிகவியல்': 'physics',
  'வேதியியல்': 'chemistry',
  'இரசாயனவியல்': 'chemistry',
  'உயிரியல்': 'biology',
  'கணிதம்': 'mathematics',
  'விஞ்ஞானம்': 'science',
  'வரலாறு': 'history',
  'புவியியல்': 'geography',
  'தமிழ்': 'tamil',
  'ஆங்கிலம்': 'english',
  'சமயம்': 'religion',
  'சித்திரம்': 'art',
  'சங்கீதம்': 'music',
  'இசை': 'music',
  'விவசாயம்': 'agriculture',
  'சுற்றாடல்': 'environment',
  'சுகாதாரம்': 'health',
  'புகைப்படவியல்': 'photography',
  'காணொளி': 'video',
  'விடுதி': 'hostel',

  // Accounting & Financial
  'வரவு': 'credit_income',
  'செலவு': 'debit_expense',
  'இருப்பு': 'balance_stock',
  'மீதி': 'balance',
  'நிகர மீதி': 'net_balance',
  'ஆரம்ப மீதி': 'opening_balance',
  'இறுதி மீதி': 'closing_balance',
  'ரசீது இலக்கம்': 'receipt_no',
  'ரசீது எண்': 'receipt_no',
  'ரசீது': 'receipt',
  'நன்கொடை': 'donation',
  'நன்கொடையாளர்': 'donor_name',
  'நன்கொடையாளர் பெயர்': 'donor_name',
  'பங்களிப்பு': 'contribution',
  'நிதி': 'fund',
  'வங்கி': 'bank',
  'காசோலை எண்': 'cheque_no',
  'காசோலை': 'cheque',
  'பணம்': 'cash',
  'கட்டணம்': 'fee',
  'அங்கீகரித்தவர்': 'approved_by',
  'அனுமதித்தவர்': 'approved_by',
  'சரிபார்த்தவர்': 'verified_by',
  'தயாரித்தவர்': 'prepared_by',

  // Students & Academics
  'மாணவர் பெயர்': 'student_name',
  'மாணவர் சுட்டெண்': 'student_id',
  'மாணவர் இலக்கம்': 'student_id',
  'சுட்டெண்': 'index_no',
  'அனுமதி எண்': 'admission_no',
  'பதிவு எண்': 'registration_no',
  'பெயர்': 'name',
  'ஆசிரியர் பெயர்': 'teacher_name',
  'ஆசிரியர்': 'teacher',
  'பாடம்': 'subject',
  'பாடங்கள்': 'subjects',
  'புள்ளி': 'marks',
  'புள்ளிகள்': 'marks',
  'தரம்': 'grade_rank',
  'நிலை': 'status',
  'வருகை': 'attendance',
  'பிரசன்னம்': 'present',
  'சமூகம்': 'present',
  'வராதவர்': 'absent',
  'பிந்தி': 'late',
  'அனுமதி': 'permission',
  'பிறந்த திகதி': 'date_of_birth',
  'பிறந்த தேதி': 'date_of_birth',
  'வயது': 'age',
  'பால்': 'gender',
  'ஆண்': 'male',
  'பெண்': 'female',
  'இரத்த வகை': 'blood_group',
  'முகவரி': 'address',
  'ஊர்': 'city_town',
  'மாவட்டம்': 'district',
  'தேசிய அடையாள அட்டை எண்': 'nic_no',
  'அடையாள அட்டை': 'id_card',
  'ஆண்டு': 'academic_year',
  'வருடம்': 'year',
  'திகதி': 'date',
  'தேதி': 'date',
  'மாதம்': 'month',
  'பருவம்': 'term',
  'தவணை': 'term'
};

// Character transliteration map for phonetic conversion of unmatched Tamil syllables
const TAMIL_CHAR_MAP: Record<string, string> = {
  'அ': 'a', 'ஆ': 'aa', 'இ': 'i', 'ஈ': 'ee', 'உ': 'u', 'ஊ': 'oo',
  'எ': 'e', 'ஏ': 'ae', 'ஐ': 'ai', 'ஒ': 'o', 'ஓ': 'oa', 'ஔ': 'au',
  'ஃ': 'h',
  'க': 'ka', 'ங': 'nga', 'ச': 'sa', 'ஞ': 'nya', 'ட': 'ta', 'ண': 'na',
  'த': 'tha', 'ந': 'na', 'ப': 'pa', 'ம': 'ma', 'ய': 'ya', 'ர': 'ra',
  'ல': 'la', 'வ': 'va', 'ழ': 'zha', 'ள': 'la', 'ற': 'ra', 'ன': 'na',
  'ஜ': 'ja', 'ஷ': 'sha', 'ஸ': 'sa', 'ஹ': 'ha', 'க்ஷ': 'ksha',
  '்': '', 'ா': 'aa', 'ி': 'i', 'ீ': 'ee', 'ு': 'u', 'ூ': 'oo',
  'ெ': 'e', 'ே': 'ae', 'ை': 'ai', 'ொ': 'o', 'ோ': 'oa', 'ௌ': 'au'
};

/**
 * Checks if a string contains any Tamil unicode characters (\u0B80 - \u0BFF)
 */
export function containsTamil(text: string): boolean {
  if (!text) return false;
  return /[\u0B80-\u0BFF]/.test(text);
}

/**
 * Phonetically transliterates Tamil text into readable English characters
 */
export function transliterateTamilToEnglish(text: string): string {
  if (!text) return '';
  let result = '';
  const len = text.length;
  
  for (let i = 0; i < len; i++) {
    const char = text[i];
    if (TAMIL_CHAR_MAP[char] !== undefined) {
      // Check if next character is a vowel sign (புள்ளி or மாத்திரை)
      const nextChar = i + 1 < len ? text[i + 1] : '';
      if (nextChar === '்') {
        // Pure consonant without inherent 'a'
        const base = TAMIL_CHAR_MAP[char].replace(/a$/, '');
        result += base;
        i++; // skip pulli
      } else if (['ா', 'ி', 'ீ', 'ு', 'ூ', 'ெ', 'ே', 'ை', 'ொ', 'ோ', 'ௌ'].includes(nextChar)) {
        const base = TAMIL_CHAR_MAP[char].replace(/a$/, '');
        const vowel = TAMIL_CHAR_MAP[nextChar];
        result += base + vowel;
        i++; // skip vowel sign
      } else {
        result += TAMIL_CHAR_MAP[char];
      }
    } else if (/[a-zA-Z0-9_]/.test(char)) {
      result += char;
    } else if (/\s+/.test(char)) {
      result += '_';
    }
  }
  
  return result.replace(/_+/g, '_').replace(/^_+|_+$/g, '');
}

/**
 * Translates a Tamil header or mixed header into a clean, intuitive English snake_case database column name.
 */
export function translateTamilHeader(raw: string): string {
  if (!raw) return 'col';
  const cleanTrimmed = raw.trim();

  // 1. Direct exact dictionary match
  if (TAMIL_DICTIONARY[cleanTrimmed]) {
    return TAMIL_DICTIONARY[cleanTrimmed];
  }

  // 2. Normalized dictionary match (strip punctuation, whitespace)
  const normalizedKey = cleanTrimmed
    .replace(/[.,/#!$%^&*;:{}=\-_`~()]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (TAMIL_DICTIONARY[normalizedKey]) {
    return TAMIL_DICTIONARY[normalizedKey];
  }

  // 3. Substring & multi-word phrase matching
  let processed = normalizedKey;
  let translatedAny = false;

  // Sort dictionary keys by descending length to match longest multi-word phrases first
  const sortedKeys = Object.keys(TAMIL_DICTIONARY).sort((a, b) => b.length - a.length);

  for (const key of sortedKeys) {
    if (processed.includes(key)) {
      processed = processed.replace(new RegExp(key, 'g'), ` ${TAMIL_DICTIONARY[key]} `);
      translatedAny = true;
    }
  }

  if (translatedAny) {
    // If some Tamil words were translated, clean up the result
    return processed
      .toLowerCase()
      .replace(/[^a-z0-9_]/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_+|_+$/g, '') || 'col_data';
  }

  // 4. If text contains Tamil but was not in dictionary, phonetically transliterate it
  if (containsTamil(raw)) {
    const transliterated = transliterateTamilToEnglish(raw);
    if (transliterated && transliterated.length >= 2) {
      return transliterated.toLowerCase();
    }
  }

  // 5. Normal English fallback
  return raw
    .toLowerCase()
    .replace(/[^a-z0-9_]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '') || 'col';
}

/**
 * Universal Intelligent Column & Table Identifier Sanitizer.
 * Automatically detects Tamil headers, translates or transliterates them,
 * and outputs clean, SQL-safe snake_case database identifiers.
 */
export function smartSanitizeIdentifier(raw: string, fallbackPrefix: string = 'col'): string {
  if (!raw || !raw.trim()) return fallbackPrefix;
  const trimmed = raw.trim();

  // If text contains Tamil characters, translate or transliterate first
  if (containsTamil(trimmed)) {
    const translated = translateTamilHeader(trimmed);
    if (translated && translated !== 'col' && translated !== 'col_data') {
      return translated;
    }
  }

  // Standard English sanitization
  let s = trimmed.toLowerCase();
  s = s.replace(/[^a-z0-9_]/g, '_');
  s = s.replace(/_+/g, '_');
  s = s.replace(/^_+|_+$/g, '');

  if (!s) {
    return fallbackPrefix;
  }

  if (/^[0-9]/.test(s)) {
    s = `${fallbackPrefix}_${s}`;
  }

  // Reserved Postgres words
  const reserved = [
    'user', 'order', 'group', 'table', 'select', 'where', 'limit', 'offset',
    'primary', 'check', 'index', 'column', 'values', 'database', 'from', 'into', 'join'
  ];
  if (reserved.includes(s)) {
    s = `${s}_val`;
  }

  return s || fallbackPrefix;
}
