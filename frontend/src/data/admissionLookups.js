import * as XLSX from 'xlsx';

// Official Kannur University Lookups & Standard Target Schema
// Sourced from kreap26vfyfyimp2026updated_working file.xlsx (Kannur IDs sheet)

export const KANNUR_CATEGORIES = {
  "SEBC-Other Backward Hindu": "10",
  "SEBC-Muslim": "10",
  "SEBC-Ezhava/Thiyya/Billava": "10",
  "Scheduled Castes": "2",
  "General": "1",
  "EWS": "1",
  "Scheduled Tribes": "3",
  "SEBC-Latin Catholic(LC)": "10"
};

export const CATEGORY_LOOKUP = {
  "sebcotherbackwardhindu": { id: "10", name: "SEBC-Other Backward Hindu" },
  "sebc-other backward hindu": { id: "10", name: "SEBC-Other Backward Hindu" },
  "sebcmuslim": { id: "10", name: "SEBC-Muslim" },
  "sebc-muslim": { id: "10", name: "SEBC-Muslim" },
  "sebcezhavathiyyabillava": { id: "10", name: "SEBC-Ezhava/Thiyya/Billava" },
  "sebc-ezhava/thiyya/billava": { id: "10", name: "SEBC-Ezhava/Thiyya/Billava" },
  "scheduledcastes": { id: "2", name: "Scheduled Castes" },
  "scheduled castes": { id: "2", name: "Scheduled Castes" },
  "general": { id: "1", name: "General" },
  "general": { id: "1", name: "General" },
  "ews": { id: "1", name: "EWS" },
  "ews": { id: "1", name: "EWS" },
  "scheduledtribes": { id: "3", name: "Scheduled Tribes" },
  "scheduled tribes": { id: "3", name: "Scheduled Tribes" },
  "sebclatincatholiclc": { id: "10", name: "SEBC-Latin Catholic(LC)" },
  "sebc-latin catholic(lc)": { id: "10", name: "SEBC-Latin Catholic(LC)" },
  "general": { id: "1", name: "General" },
  "gen": { id: "1", name: "General" },
  "open": { id: "1", name: "General" },
  "ews": { id: "1", name: "EWS" },
  "sc": { id: "2", name: "Scheduled Castes" },
  "scheduled caste": { id: "2", name: "Scheduled Castes" },
  "scheduled castes": { id: "2", name: "Scheduled Castes" },
  "st": { id: "3", name: "Scheduled Tribes" },
  "scheduled tribe": { id: "3", name: "Scheduled Tribes" },
  "scheduled tribes": { id: "3", name: "Scheduled Tribes" },
  "sebc": { id: "10", name: "SEBC" },
  "obc": { id: "10", name: "SEBC" },
  "general/oec": { id: "10", name: "General/OEC" },
  "generaloec": { id: "10", name: "General/OEC" },
  "oec": { id: "10", name: "General/OEC" },
  "sebc-obh": { id: "10", name: "SEBC-Other Backward Hindu" },
  "sebcobh": { id: "10", name: "SEBC-Other Backward Hindu" },
  "sebc-other backward christian": { id: "10", name: "SEBC-Other Backward Christian" },
  "sebcotherbackwardchristian": { id: "10", name: "SEBC-Other Backward Christian" },
  "sebc-other backward hindu": { id: "10", name: "SEBC-Other Backward Hindu" },
  "sebcotherbackwardhindu": { id: "10", name: "SEBC-Other Backward Hindu" },
  "sebc-muslim": { id: "10", name: "SEBC-Muslim" },
  "sebcmuslim": { id: "10", name: "SEBC-Muslim" },
  "sebc-ezhava/thiyya/billava": { id: "10", name: "SEBC-Ezhava/Thiyya/Billava" },
  "sebcezhavathiyyabillava": { id: "10", name: "SEBC-Ezhava/Thiyya/Billava" },
  "muslim": { id: "10", name: "SEBC-Muslim" },
  "thiyya": { id: "10", name: "SEBC-Ezhava/Thiyya/Billava" },
  "thiyyas": { id: "10", name: "SEBC-Ezhava/Thiyya/Billava" },
  "ezhava": { id: "10", name: "SEBC-Ezhava/Thiyya/Billava" },
  "latin": { id: "10", name: "SEBC-Latin Catholic(LC)" },
  "latin catholic": { id: "10", name: "SEBC-Latin Catholic(LC)" }
};

export const KANNUR_RELIGIONS = {
  "Hindu": "1",
  "Islam": "2",
  "Christian": "3",
  "Others": "5"
};

export const RELIGION_LOOKUP = {
  "hindu": { id: "1", name: "Hindu" },
  "hindu": { id: "1", name: "Hindu" },
  "islam": { id: "2", name: "Islam" },
  "islam": { id: "2", name: "Islam" },
  "christian": { id: "3", name: "Christian" },
  "christian": { id: "3", name: "Christian" },
  "others": { id: "5", name: "Others" },
  "others": { id: "5", name: "Others" },
  "hindu": { id: "1", name: "Hindu" },
  "islam": { id: "2", name: "Islam" },
  "muslim": { id: "2", name: "Islam" },
  "christian": { id: "3", name: "Christian" },
  "others": { id: "5", name: "Others" }
};

export const CASTE_LOOKUP = {
  "adidravida": "178",
  "adi dravida": "178",
  "Adi Dravida": "178",
  "adikarnataka": "223",
  "adi karnataka": "223",
  "Adi Karnataka": "223",
  "adiyan": "328",
  "adiyan": "328",
  "Adiyan": "328",
  "agasa": "2",
  "agasa": "2",
  "Agasa": "2",
  "andhurunair": "183",
  "andhuru nair": "183",
  "Andhuru Nair": "183",
  "arayan": "6",
  "arayan": "6",
  "Arayan": "6",
  "arunthathiyar": "225",
  "arunthathiyar": "225",
  "Arunthathiyar": "225",
  "arya": "10",
  "arya": "10",
  "Arya": "10",
  "asari": "11",
  "asari": "11",
  "Asari": "11",
  "baira": "227",
  "baira": "227",
  "Baira": "227",
  "bakuda": "228",
  "bakuda": "228",
  "Bakuda": "228",
  "bestha": "15",
  "bestha": "15",
  "Bestha": "15",
  "bhandariorbhondari": "16",
  "bhandari or bhondari": "16",
  "Bhandari or Bhondari": "16",
  "bharatharotherthanparathar": "230",
  "bharathar(other than parathar)": "230",
  "Bharathar(Other than Parathar)": "230",
  "billava": "312",
  "billava": "312",
  "Billava": "312",
  "bovies": "17",
  "bovies": "17",
  "Bovies": "17",
  "chakkiliyan": "231",
  "chakkiliyan": "231",
  "Chakkiliyan": "231",
  "chaliya": "22",
  "chaliya": "22",
  "Chaliya": "22",
  "cheramar": "238",
  "cheramar": "238",
  "Cheramar": "238",
  "cheruman": "239",
  "cheruman": "239",
  "Cheruman": "239",
  "chettychetties": "313",
  "chetty/chetties": "313",
  "Chetty/Chetties": "313",
  "convertsfromscheduledcastestochristianity": "314",
  "converts from scheduled castes to christianity": "314",
  "Converts from Scheduled Castes to Christianity": "314",
  "devadiga": "26",
  "devadiga": "26",
  "Devadiga": "26",
  "devanga": "27",
  "devanga": "27",
  "Devanga": "27",
  "dheevara": "28",
  "dheevara": "28",
  "Dheevara": "28",
  "eruman": "31",
  "eruman": "31",
  "Eruman": "31",
  "ezhavas": "316",
  "ezhavas": "316",
  "Ezhavas": "316",
  "ezhuthachan": "33",
  "ezhuthachan": "33",
  "Ezhuthachan": "33",
  "ganaka": "35",
  "ganaka": "35",
  "Ganaka": "35",
  "gatti": "38",
  "gatti": "38",
  "Gatti": "38",
  "general": "393",
  "general": "393",
  "General": "393",
  "gowda": "40",
  "gowda": "40",
  "Gowda": "40",
  "hegde": "42",
  "hegde": "42",
  "Hegde": "42",
  "hindunadar": "43",
  "hindu nadar": "43",
  "Hindu Nadar": "43",
  "irivavan": "192",
  "irivavan": "192",
  "Irivavan": "192",
  "irular": "334",
  "irular": "334",
  "Irular": "334",
  "islam": "323",
  "islam": "323",
  "Islam": "323",
  "jogi": "50",
  "jogi": "50",
  "Jogi": "50",
  "kadarwayanadkadar": "336",
  "kadar (wayanad kadar)": "336",
  "Kadar (Wayanad Kadar)": "336",
  "kalarikuruporkalaripanicker": "439",
  "kalarikurup or kalari panicker": "439",
  "Kalarikurup or Kalari Panicker": "439",
  "kalladi": "247",
  "kalladi": "247",
  "Kalladi": "247",
  "kammalas": "59",
  "kammalas": "59",
  "Kammalas": "59",
  "kanakkan": "248",
  "kanakkan": "248",
  "Kanakkan": "248",
  "kanisan": "63",
  "kanisan": "63",
  "Kanisan": "63",
  "kaniyan": "65",
  "kaniyan": "65",
  "Kaniyan": "65",
  "kannian": "69",
  "kannian": "69",
  "Kannian": "69",
  "karimpalan": "339",
  "karimpalan": "339",
  "Karimpalan": "339",
  "karuvan": "72",
  "karuvan": "72",
  "Karuvan": "72",
  "kattunayakan": "340",
  "kattunayakan": "340",
  "Kattunayakan": "340",
  "kavara": "249",
  "kavara": "249",
  "Kavara": "249",
  "kavuthiyan": "74",
  "kavuthiyan": "74",
  "Kavuthiyan": "74",
  "kitaran": "79",
  "kitaran": "79",
  "Kitaran": "79",
  "kollan": "81",
  "kollan": "81",
  "Kollan": "81",
  "koraga": "343",
  "koraga": "343",
  "Koraga": "343",
  "koteyar": "320",
  "koteyar": "320",
  "Koteyar": "320",
  "kottarchetties": "321",
  "kottar chetties": "321",
  "Kottar Chetties": "321",
  "kudiya": "344",
  "kudiya": "344",
  "Kudiya": "344",
  "kudumbi": "85",
  "kudumbi": "85",
  "Kudumbi": "85",
  "kulala": "86",
  "kulala": "86",
  "Kulala": "86",
  "kulalan": "196",
  "kulalan": "196",
  "Kulalan": "196",
  "kurava": "254",
  "kurava": "254",
  "Kurava": "254",
  "kuravan": "255",
  "kuravan": "255",
  "Kuravan": "255",
  "kurichchankurichiyan": "347",
  "kurichchan (kurichiyan)": "347",
  "Kurichchan (Kurichiyan)": "347",
  "kuruman": "348",
  "kuruman": "348",
  "Kuruman": "348",
  "kurumans": "349",
  "kurumans": "349",
  "Kurumans": "349",
  "kusavan": "90",
  "kusavan": "90",
  "Kusavan": "90",
  "latincatholics": "322",
  "latin catholics": "322",
  "Latin Catholics": "322",
  "maila": "257",
  "maila": "257",
  "Maila": "257",
  "malavettuvan": "358",
  "mala vettuvan": "358",
  "Mala Vettuvan": "358",
  "malayan": "205",
  "malayan": "205",
  "Malayan": "205",
  "malayarayar": "365",
  "malayarayar": "365",
  "Malayarayar": "365",
  "maniyani": "95",
  "maniyani": "95",
  "Maniyani": "95",
  "mannan": "259",
  "mannan": "259",
  "Mannan": "259",
  "mappila": "323",
  "mappila": "323",
  "Mappila": "323",
  "maratisofhosdurgtaluk": "96",
  "maratis of hosdurgtaluk": "96",
  "Maratis of HosdurgTaluk": "96",
  "maratisofkasaragodandhosdurgtaluk": "367",
  "maratis of kasaragod and hosdurg taluk": "367",
  "Maratis of Kasaragod and hosdurg Taluk": "367",
  "maravanmaravar": "97",
  "maravan (maravar)": "97",
  "Maravan (Maravar)": "97",
  "maruthuvar": "98",
  "maruthuvar": "98",
  "Maruthuvar": "98",
  "mavilan": "368",
  "mavilan": "368",
  "Mavilan": "368",
  "melakudi": "369",
  "melakudi": "369",
  "Melakudi": "369",
  "mogerotherthanmogeyar": "261",
  "moger (other than mogeyar)": "261",
  "Moger (Other than Mogeyar)": "261",
  "mokaya": "208",
  "mokaya": "208",
  "Mokaya": "208",
  "moosari": "108",
  "moosari": "108",
  "Moosari": "108",
  "mukaya": "109",
  "mukaya": "109",
  "Mukaya": "109",
  "mukhari": "110",
  "mukhari": "110",
  "Mukhari": "110",
  "mukkuvan": "111",
  "mukkuvan": "111",
  "Mukkuvan": "111",
  "mundala": "263",
  "mundala": "263",
  "Mundala": "263",
  "muslims": "323",
  "muslims": "323",
  "Muslims": "323",
  "muvari": "112",
  "muvari": "112",
  "Muvari": "112",
  "nadarbelongingtochristianreligiousdenominationsotherthansiuc": "324",
  "nadar belonging to christian religious denominations other than siuc": "324",
  "Nadar belonging to Christian religious denominations other than SIUC": "324",
  "naicken": "113",
  "naicken": "113",
  "Naicken": "113",
  "nalkadaya": "265",
  "nalkadaya": "265",
  "Nalkadaya": "265",
  "notmentioned": "393",
  "not mentioned": "393",
  "Not Mentioned": "393",
  "other": "268",
  "other": "268",
  "Other": "268",
  "padmasali": "120",
  "padmasali": "120",
  "Padmasali": "120",
  "pallan": "223",
  "pallan": "223",
  "Pallan": "223",
  "paniyan": "383",
  "paniyan": "383",
  "Paniyan": "383",
  "paravan": "277",
  "paravan": "277",
  "Paravan": "277",
  "paraya": "278",
  "paraya": "278",
  "Paraya": "278",
  "parayan": "279",
  "parayan": "279",
  "Parayan": "279",
  "pathiyan": "281",
  "pathiyan": "281",
  "Pathiyan": "281",
  "perumannan": "282",
  "perumannan": "282",
  "Perumannan": "282",
  "perumkollen": "129",
  "perumkollen": "129",
  "Perumkollen": "129",
  "pulaya": "284",
  "pulaya": "284",
  "Pulaya": "284",
  "pulayan": "286",
  "pulayan": "286",
  "Pulayan": "286",
  "pulluvan": "288",
  "pulluvan": "288",
  "Pulluvan": "288",
  "rajapuri": "131",
  "rajapuri": "131",
  "Rajapuri": "131",
  "siuc": "326",
  "siuc": "326",
  "SIUC": "326",
  "sadhuchetty": "132",
  "sadhu chetty": "132",
  "Sadhu Chetty": "132",
  "saliyas": "136",
  "saliyas": "136",
  "Saliyas": "136",
  "sambava": "293",
  "sambava": "293",
  "Sambava": "293",
  "scheduledtribeslakshadweep": "393",
  "scheduled tribes (lakshadweep)": "393",
  "Scheduled Tribes (Lakshadweep)": "393",
  "sidhanar": "298",
  "sidhanar": "298",
  "Sidhanar": "298",
  "thandan": "301",
  "thandan": "301",
  "Thandan": "301",
  "thattan": "144",
  "thattan": "144",
  "Thattan": "144",
  "thiyyas": "327",
  "thiyyas": "327",
  "Thiyyas": "327",
  "thiyyaswhoareknownasthandanintheerstwhilecochinandmalabarareas": "327",
  "thiyyas (who are known as thandan, in the erstwhile cochin and malabar areas)": "327",
  "Thiyyas (who are known as Thandan, in the erstwhile Cochin and Malabar areas)": "327",
  "uraly": "391",
  "uraly": "391",
  "Uraly": "391",
  "vakkaliga": "155",
  "vakkaliga": "155",
  "Vakkaliga": "155",
  "valan": "157",
  "valan": "157",
  "Valan": "157",
  "valluvan": "304",
  "valluvan": "304",
  "Valluvan": "304",
  "vanikavaisya": "160",
  "vanika vaisya": "160",
  "Vanika Vaisya": "160",
  "vaniya": "161",
  "vaniya": "161",
  "Vaniya": "161",
  "vaniyan": "162",
  "vaniyan": "162",
  "Vaniyan": "162",
  "vannan": "305",
  "vannan": "305",
  "Vannan": "305",
  "vannathan": "163",
  "vannathan": "163",
  "Vannathan": "163",
  "veerasaivas": "164",
  "veera saivas": "164",
  "Veera Saivas": "164",
  "velan": "306",
  "velan": "306",
  "Velan": "306",
  "veluthedan": "166",
  "veluthedan": "166",
  "Veluthedan": "166",
  "veluthedathunair": "167",
  "veluthedathu nair": "167",
  "Veluthedathu Nair": "167",
  "vettakuruman": "392",
  "vetta kuruman": "392",
  "Vetta Kuruman": "392",
  "vettuvan": "308",
  "vettuvan": "308",
  "Vettuvan": "308",
  "vilakkathalavan": "169",
  "vilakkathalavan": "169",
  "Vilakkathalavan": "169",
  "vilakkithalanair": "170",
  "vilakkithala nair": "170",
  "Vilakkithala Nair": "170",
  "viswakarmas": "175",
  "viswakarmas": "175",
  "Viswakarmas": "175",
  "wynadanchetty": "176",
  "wynadanchetty": "176",
  "WynadanChetty": "176",
  "yadava": "177",
  "yadava": "177",
  "Yadava": "177",
  "yogi": "180",
  "yogi": "180",
  "Yogi": "180",
  "general": "1",
  "nair": "1",
  "muslim": "168",
  "mappila": "168",
  "thiyya": "169",
  "thiyyas": "169",
  "ezhava": "169",
  "mukkuvan": "320"
};

export const STREAM_LOOKUP = {
  "commerce": { id: "2", name: "Commerce" },
  "commerce": { id: "2", name: "Commerce" },
  "humanities": { id: "3", name: "Humanities" },
  "humanities": { id: "3", name: "Humanities" },
  "science": { id: "1", name: "Science" },
  "science": { id: "1", name: "Science" },
};

export const QUALIFICATION_LOOKUP = {
  "hsekerala": { facultyId: "2", name: "HSE - Kerala" },
  "hse - kerala": { facultyId: "2", name: "HSE - Kerala" },
  "cbse": { facultyId: "2", name: "CBSE" },
  "cbse": { facultyId: "2", name: "CBSE" },
  "vhsekeralahsensqf": { facultyId: "2", name: "VHSE - Kerala/HSE NSQF" },
  "vhse - kerala/hse nsqf": { facultyId: "2", name: "VHSE - Kerala/HSE NSQF" },
  "others": { facultyId: "7", name: "Others" },
  "others": { facultyId: "7", name: "Others" },
  "isc": { facultyId: "2", name: "ISC" },
  "isc": { facultyId: "2", name: "ISC" },
};

export const KANNUR_PROGRAMMES = [
  {
    "pgmName": "Five Year Integrated Masters Programme in  Clinical Psychology",
    "code": "PFPSYGS25",
    "workflowId": "151",
    "name": "Five Year Integrated Masters Programme in Clinical Psychology(with Credits)-Regular-FYIMP-2025-Clinical Psychology-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in Commerce",
    "code": "PFCOMGS25",
    "workflowId": "148",
    "name": "Five Year Integrated Masters Programme in Commerce(with Credits)-Regular-FYIMP-2025-Commerce-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in Computational science",
    "code": "PFCLSGS25",
    "workflowId": "149",
    "name": "Five Year Integrated Masters Programme in Computational science(with Credits)-Regular-FYIMP-2025-Computational Sciences-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in  Geography",
    "code": "PFGEOGS25",
    "workflowId": "152",
    "name": "Five Year Integrated Masters Programme in Geography(with Credits)-Regular-FYIMP-2025-Geography-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in Music",
    "code": "PFMUSGA25",
    "workflowId": "153",
    "name": "Five Year Integrated Masters Programme in Music(with Credits)-Regular-FYIMP-2025-Music-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in Physical Education and Sports",
    "code": "PFPESGS25",
    "workflowId": "150",
    "name": "Five Year Integrated Masters Programme in Physical Education and Sports (IMPES)(with Credits)-Regular-FYIMP-2025-Physical Education and Sports-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in  Business Administration",
    "code": "PFMBAGA25",
    "workflowId": "162",
    "name": "Five Year Integrated Masters Programme in Business Administration(with Credits)-Regular-FYIMP-2025-Business Administration-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in Anthropological sciences",
    "code": "PFANTGS25",
    "workflowId": "154",
    "name": "Five Year Integrated Masters Programme in Anthropological sciences(with Credits)-Regular-FYIMP-2025-Anthropological Sciences-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in English Studies",
    "code": "PFENGGA25",
    "workflowId": "155",
    "name": "Five Year Integrated Masters Programme in English Studies(with Credits)-Regular-FYIMP-2025-English Studies-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in Envirnmental Sciences",
    "code": "PFEVSGS25",
    "workflowId": "156",
    "name": "Five Year Integrated Masters Programme in Envirnmental Sciences(with Credits)-Regular-FYIMP-2025-Environmental Science-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in Historical Studies",
    "code": "PFHISGS25",
    "workflowId": "160",
    "name": "Five Year Integrated Masters Programme in Historical Studies(with Credits)-Regular-FYIMP-2025-Historical Studies-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in Mathematical Sciences",
    "code": "PFMATGS25",
    "workflowId": "160",
    "name": "Five Year Integrated Masters Programme in Historical Studies(with Credits)-Regular-FYIMP-2025-Historical Studies-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in Physical Science",
    "code": "PFPLSGS25",
    "workflowId": "159",
    "name": "Five Year Integrated Masters Programme in Physical Science(with Credits)-Regular-FYIMP-2025-Physical Science-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Masters Programme in Social Science with Specialization  in Economics",
    "code": "PFSSEGS25",
    "workflowId": "161",
    "name": "Five Year Integrated Masters Programme in Social Science with Specialization  in Economics(with Credits)-Regular-FYIMP-2025-Economics-FYIMP Year I SEMESTER I"
  },
  {
    "pgmName": "Five Year Integrated Programe in Master of Computer Application",
    "code": "PFMCAGS25",
    "workflowId": "158",
    "name": "Five Year Integrated Programe in Master of Computer Application(with Credits)-Regular-FYIMP-2025-Computer Application-FYIMP Year I SEMESTER I"
  }
];

export const PROGRAM_LOOKUP = {
  "fiveyearintegratedmastersprogrammeinclinicalpsychology": {
    "pgmName": "Five Year Integrated Masters Programme in  Clinical Psychology",
    "code": "PFPSYGS25",
    "workflowId": "151",
    "name": "Five Year Integrated Masters Programme in Clinical Psychology(with Credits)-Regular-FYIMP-2025-Clinical Psychology-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in  clinical psychology": {
    "pgmName": "Five Year Integrated Masters Programme in  Clinical Psychology",
    "code": "PFPSYGS25",
    "workflowId": "151",
    "name": "Five Year Integrated Masters Programme in Clinical Psychology(with Credits)-Regular-FYIMP-2025-Clinical Psychology-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinclinicalpsychologywithcreditsregularfyimp2025clinicalpsychologyfyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in  Clinical Psychology",
    "code": "PFPSYGS25",
    "workflowId": "151",
    "name": "Five Year Integrated Masters Programme in Clinical Psychology(with Credits)-Regular-FYIMP-2025-Clinical Psychology-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in clinical psychology(with credits)-regular-fyimp-2025-clinical psychology-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in  Clinical Psychology",
    "code": "PFPSYGS25",
    "workflowId": "151",
    "name": "Five Year Integrated Masters Programme in Clinical Psychology(with Credits)-Regular-FYIMP-2025-Clinical Psychology-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeincommerce": {
    "pgmName": "Five Year Integrated Masters Programme in Commerce",
    "code": "PFCOMGS25",
    "workflowId": "148",
    "name": "Five Year Integrated Masters Programme in Commerce(with Credits)-Regular-FYIMP-2025-Commerce-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in commerce": {
    "pgmName": "Five Year Integrated Masters Programme in Commerce",
    "code": "PFCOMGS25",
    "workflowId": "148",
    "name": "Five Year Integrated Masters Programme in Commerce(with Credits)-Regular-FYIMP-2025-Commerce-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeincommercewithcreditsregularfyimp2025commercefyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in Commerce",
    "code": "PFCOMGS25",
    "workflowId": "148",
    "name": "Five Year Integrated Masters Programme in Commerce(with Credits)-Regular-FYIMP-2025-Commerce-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in commerce(with credits)-regular-fyimp-2025-commerce-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in Commerce",
    "code": "PFCOMGS25",
    "workflowId": "148",
    "name": "Five Year Integrated Masters Programme in Commerce(with Credits)-Regular-FYIMP-2025-Commerce-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeincomputationalscience": {
    "pgmName": "Five Year Integrated Masters Programme in Computational science",
    "code": "PFCLSGS25",
    "workflowId": "149",
    "name": "Five Year Integrated Masters Programme in Computational science(with Credits)-Regular-FYIMP-2025-Computational Sciences-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in computational science": {
    "pgmName": "Five Year Integrated Masters Programme in Computational science",
    "code": "PFCLSGS25",
    "workflowId": "149",
    "name": "Five Year Integrated Masters Programme in Computational science(with Credits)-Regular-FYIMP-2025-Computational Sciences-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeincomputationalsciencewithcreditsregularfyimp2025computationalsciencesfyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in Computational science",
    "code": "PFCLSGS25",
    "workflowId": "149",
    "name": "Five Year Integrated Masters Programme in Computational science(with Credits)-Regular-FYIMP-2025-Computational Sciences-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in computational science(with credits)-regular-fyimp-2025-computational sciences-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in Computational science",
    "code": "PFCLSGS25",
    "workflowId": "149",
    "name": "Five Year Integrated Masters Programme in Computational science(with Credits)-Regular-FYIMP-2025-Computational Sciences-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeingeography": {
    "pgmName": "Five Year Integrated Masters Programme in  Geography",
    "code": "PFGEOGS25",
    "workflowId": "152",
    "name": "Five Year Integrated Masters Programme in Geography(with Credits)-Regular-FYIMP-2025-Geography-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in  geography": {
    "pgmName": "Five Year Integrated Masters Programme in  Geography",
    "code": "PFGEOGS25",
    "workflowId": "152",
    "name": "Five Year Integrated Masters Programme in Geography(with Credits)-Regular-FYIMP-2025-Geography-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeingeographywithcreditsregularfyimp2025geographyfyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in  Geography",
    "code": "PFGEOGS25",
    "workflowId": "152",
    "name": "Five Year Integrated Masters Programme in Geography(with Credits)-Regular-FYIMP-2025-Geography-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in geography(with credits)-regular-fyimp-2025-geography-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in  Geography",
    "code": "PFGEOGS25",
    "workflowId": "152",
    "name": "Five Year Integrated Masters Programme in Geography(with Credits)-Regular-FYIMP-2025-Geography-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinmusic": {
    "pgmName": "Five Year Integrated Masters Programme in Music",
    "code": "PFMUSGA25",
    "workflowId": "153",
    "name": "Five Year Integrated Masters Programme in Music(with Credits)-Regular-FYIMP-2025-Music-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in music": {
    "pgmName": "Five Year Integrated Masters Programme in Music",
    "code": "PFMUSGA25",
    "workflowId": "153",
    "name": "Five Year Integrated Masters Programme in Music(with Credits)-Regular-FYIMP-2025-Music-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinmusicwithcreditsregularfyimp2025musicfyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in Music",
    "code": "PFMUSGA25",
    "workflowId": "153",
    "name": "Five Year Integrated Masters Programme in Music(with Credits)-Regular-FYIMP-2025-Music-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in music(with credits)-regular-fyimp-2025-music-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in Music",
    "code": "PFMUSGA25",
    "workflowId": "153",
    "name": "Five Year Integrated Masters Programme in Music(with Credits)-Regular-FYIMP-2025-Music-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinphysicaleducationandsports": {
    "pgmName": "Five Year Integrated Masters Programme in Physical Education and Sports",
    "code": "PFPESGS25",
    "workflowId": "150",
    "name": "Five Year Integrated Masters Programme in Physical Education and Sports (IMPES)(with Credits)-Regular-FYIMP-2025-Physical Education and Sports-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in physical education and sports": {
    "pgmName": "Five Year Integrated Masters Programme in Physical Education and Sports",
    "code": "PFPESGS25",
    "workflowId": "150",
    "name": "Five Year Integrated Masters Programme in Physical Education and Sports (IMPES)(with Credits)-Regular-FYIMP-2025-Physical Education and Sports-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinphysicaleducationandsportsimpeswithcreditsregularfyimp2025physicaleducationandsportsfyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in Physical Education and Sports",
    "code": "PFPESGS25",
    "workflowId": "150",
    "name": "Five Year Integrated Masters Programme in Physical Education and Sports (IMPES)(with Credits)-Regular-FYIMP-2025-Physical Education and Sports-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in physical education and sports (impes)(with credits)-regular-fyimp-2025-physical education and sports-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in Physical Education and Sports",
    "code": "PFPESGS25",
    "workflowId": "150",
    "name": "Five Year Integrated Masters Programme in Physical Education and Sports (IMPES)(with Credits)-Regular-FYIMP-2025-Physical Education and Sports-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinbusinessadministration": {
    "pgmName": "Five Year Integrated Masters Programme in  Business Administration",
    "code": "PFMBAGA25",
    "workflowId": "162",
    "name": "Five Year Integrated Masters Programme in Business Administration(with Credits)-Regular-FYIMP-2025-Business Administration-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in  business administration": {
    "pgmName": "Five Year Integrated Masters Programme in  Business Administration",
    "code": "PFMBAGA25",
    "workflowId": "162",
    "name": "Five Year Integrated Masters Programme in Business Administration(with Credits)-Regular-FYIMP-2025-Business Administration-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinbusinessadministrationwithcreditsregularfyimp2025businessadministrationfyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in  Business Administration",
    "code": "PFMBAGA25",
    "workflowId": "162",
    "name": "Five Year Integrated Masters Programme in Business Administration(with Credits)-Regular-FYIMP-2025-Business Administration-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in business administration(with credits)-regular-fyimp-2025-business administration-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in  Business Administration",
    "code": "PFMBAGA25",
    "workflowId": "162",
    "name": "Five Year Integrated Masters Programme in Business Administration(with Credits)-Regular-FYIMP-2025-Business Administration-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinanthropologicalsciences": {
    "pgmName": "Five Year Integrated Masters Programme in Anthropological sciences",
    "code": "PFANTGS25",
    "workflowId": "154",
    "name": "Five Year Integrated Masters Programme in Anthropological sciences(with Credits)-Regular-FYIMP-2025-Anthropological Sciences-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in anthropological sciences": {
    "pgmName": "Five Year Integrated Masters Programme in Anthropological sciences",
    "code": "PFANTGS25",
    "workflowId": "154",
    "name": "Five Year Integrated Masters Programme in Anthropological sciences(with Credits)-Regular-FYIMP-2025-Anthropological Sciences-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinanthropologicalscienceswithcreditsregularfyimp2025anthropologicalsciencesfyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in Anthropological sciences",
    "code": "PFANTGS25",
    "workflowId": "154",
    "name": "Five Year Integrated Masters Programme in Anthropological sciences(with Credits)-Regular-FYIMP-2025-Anthropological Sciences-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in anthropological sciences(with credits)-regular-fyimp-2025-anthropological sciences-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in Anthropological sciences",
    "code": "PFANTGS25",
    "workflowId": "154",
    "name": "Five Year Integrated Masters Programme in Anthropological sciences(with Credits)-Regular-FYIMP-2025-Anthropological Sciences-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinenglishstudies": {
    "pgmName": "Five Year Integrated Masters Programme in English Studies",
    "code": "PFENGGA25",
    "workflowId": "155",
    "name": "Five Year Integrated Masters Programme in English Studies(with Credits)-Regular-FYIMP-2025-English Studies-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in english studies": {
    "pgmName": "Five Year Integrated Masters Programme in English Studies",
    "code": "PFENGGA25",
    "workflowId": "155",
    "name": "Five Year Integrated Masters Programme in English Studies(with Credits)-Regular-FYIMP-2025-English Studies-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinenglishstudieswithcreditsregularfyimp2025englishstudiesfyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in English Studies",
    "code": "PFENGGA25",
    "workflowId": "155",
    "name": "Five Year Integrated Masters Programme in English Studies(with Credits)-Regular-FYIMP-2025-English Studies-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in english studies(with credits)-regular-fyimp-2025-english studies-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in English Studies",
    "code": "PFENGGA25",
    "workflowId": "155",
    "name": "Five Year Integrated Masters Programme in English Studies(with Credits)-Regular-FYIMP-2025-English Studies-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinenvirnmentalsciences": {
    "pgmName": "Five Year Integrated Masters Programme in Envirnmental Sciences",
    "code": "PFEVSGS25",
    "workflowId": "156",
    "name": "Five Year Integrated Masters Programme in Envirnmental Sciences(with Credits)-Regular-FYIMP-2025-Environmental Science-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in envirnmental sciences": {
    "pgmName": "Five Year Integrated Masters Programme in Envirnmental Sciences",
    "code": "PFEVSGS25",
    "workflowId": "156",
    "name": "Five Year Integrated Masters Programme in Envirnmental Sciences(with Credits)-Regular-FYIMP-2025-Environmental Science-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinenvirnmentalscienceswithcreditsregularfyimp2025environmentalsciencefyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in Envirnmental Sciences",
    "code": "PFEVSGS25",
    "workflowId": "156",
    "name": "Five Year Integrated Masters Programme in Envirnmental Sciences(with Credits)-Regular-FYIMP-2025-Environmental Science-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in envirnmental sciences(with credits)-regular-fyimp-2025-environmental science-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in Envirnmental Sciences",
    "code": "PFEVSGS25",
    "workflowId": "156",
    "name": "Five Year Integrated Masters Programme in Envirnmental Sciences(with Credits)-Regular-FYIMP-2025-Environmental Science-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinhistoricalstudies": {
    "pgmName": "Five Year Integrated Masters Programme in Historical Studies",
    "code": "PFHISGS25",
    "workflowId": "160",
    "name": "Five Year Integrated Masters Programme in Historical Studies(with Credits)-Regular-FYIMP-2025-Historical Studies-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in historical studies": {
    "pgmName": "Five Year Integrated Masters Programme in Historical Studies",
    "code": "PFHISGS25",
    "workflowId": "160",
    "name": "Five Year Integrated Masters Programme in Historical Studies(with Credits)-Regular-FYIMP-2025-Historical Studies-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinhistoricalstudieswithcreditsregularfyimp2025historicalstudiesfyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in Mathematical Sciences",
    "code": "PFMATGS25",
    "workflowId": "160",
    "name": "Five Year Integrated Masters Programme in Historical Studies(with Credits)-Regular-FYIMP-2025-Historical Studies-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in historical studies(with credits)-regular-fyimp-2025-historical studies-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in Mathematical Sciences",
    "code": "PFMATGS25",
    "workflowId": "160",
    "name": "Five Year Integrated Masters Programme in Historical Studies(with Credits)-Regular-FYIMP-2025-Historical Studies-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinmathematicalsciences": {
    "pgmName": "Five Year Integrated Masters Programme in Mathematical Sciences",
    "code": "PFMATGS25",
    "workflowId": "160",
    "name": "Five Year Integrated Masters Programme in Historical Studies(with Credits)-Regular-FYIMP-2025-Historical Studies-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in mathematical sciences": {
    "pgmName": "Five Year Integrated Masters Programme in Mathematical Sciences",
    "code": "PFMATGS25",
    "workflowId": "160",
    "name": "Five Year Integrated Masters Programme in Historical Studies(with Credits)-Regular-FYIMP-2025-Historical Studies-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinphysicalscience": {
    "pgmName": "Five Year Integrated Masters Programme in Physical Science",
    "code": "PFPLSGS25",
    "workflowId": "159",
    "name": "Five Year Integrated Masters Programme in Physical Science(with Credits)-Regular-FYIMP-2025-Physical Science-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in physical science": {
    "pgmName": "Five Year Integrated Masters Programme in Physical Science",
    "code": "PFPLSGS25",
    "workflowId": "159",
    "name": "Five Year Integrated Masters Programme in Physical Science(with Credits)-Regular-FYIMP-2025-Physical Science-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinphysicalsciencewithcreditsregularfyimp2025physicalsciencefyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in Physical Science",
    "code": "PFPLSGS25",
    "workflowId": "159",
    "name": "Five Year Integrated Masters Programme in Physical Science(with Credits)-Regular-FYIMP-2025-Physical Science-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in physical science(with credits)-regular-fyimp-2025-physical science-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in Physical Science",
    "code": "PFPLSGS25",
    "workflowId": "159",
    "name": "Five Year Integrated Masters Programme in Physical Science(with Credits)-Regular-FYIMP-2025-Physical Science-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinsocialsciencewithspecializationineconomics": {
    "pgmName": "Five Year Integrated Masters Programme in Social Science with Specialization  in Economics",
    "code": "PFSSEGS25",
    "workflowId": "161",
    "name": "Five Year Integrated Masters Programme in Social Science with Specialization  in Economics(with Credits)-Regular-FYIMP-2025-Economics-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in social science with specialization  in economics": {
    "pgmName": "Five Year Integrated Masters Programme in Social Science with Specialization  in Economics",
    "code": "PFSSEGS25",
    "workflowId": "161",
    "name": "Five Year Integrated Masters Programme in Social Science with Specialization  in Economics(with Credits)-Regular-FYIMP-2025-Economics-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedmastersprogrammeinsocialsciencewithspecializationineconomicswithcreditsregularfyimp2025economicsfyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Masters Programme in Social Science with Specialization  in Economics",
    "code": "PFSSEGS25",
    "workflowId": "161",
    "name": "Five Year Integrated Masters Programme in Social Science with Specialization  in Economics(with Credits)-Regular-FYIMP-2025-Economics-FYIMP Year I SEMESTER I"
  },
  "five year integrated masters programme in social science with specialization  in economics(with credits)-regular-fyimp-2025-economics-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Masters Programme in Social Science with Specialization  in Economics",
    "code": "PFSSEGS25",
    "workflowId": "161",
    "name": "Five Year Integrated Masters Programme in Social Science with Specialization  in Economics(with Credits)-Regular-FYIMP-2025-Economics-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedprogrameinmasterofcomputerapplication": {
    "pgmName": "Five Year Integrated Programe in Master of Computer Application",
    "code": "PFMCAGS25",
    "workflowId": "158",
    "name": "Five Year Integrated Programe in Master of Computer Application(with Credits)-Regular-FYIMP-2025-Computer Application-FYIMP Year I SEMESTER I"
  },
  "five year integrated programe in master of computer application": {
    "pgmName": "Five Year Integrated Programe in Master of Computer Application",
    "code": "PFMCAGS25",
    "workflowId": "158",
    "name": "Five Year Integrated Programe in Master of Computer Application(with Credits)-Regular-FYIMP-2025-Computer Application-FYIMP Year I SEMESTER I"
  },
  "fiveyearintegratedprogrameinmasterofcomputerapplicationwithcreditsregularfyimp2025computerapplicationfyimpyearisemesteri": {
    "pgmName": "Five Year Integrated Programe in Master of Computer Application",
    "code": "PFMCAGS25",
    "workflowId": "158",
    "name": "Five Year Integrated Programe in Master of Computer Application(with Credits)-Regular-FYIMP-2025-Computer Application-FYIMP Year I SEMESTER I"
  },
  "five year integrated programe in master of computer application(with credits)-regular-fyimp-2025-computer application-fyimp year i semester i": {
    "pgmName": "Five Year Integrated Programe in Master of Computer Application",
    "code": "PFMCAGS25",
    "workflowId": "158",
    "name": "Five Year Integrated Programe in Master of Computer Application(with Credits)-Regular-FYIMP-2025-Computer Application-FYIMP Year I SEMESTER I"
  },
  "bbr": {
    "code": "PFMBAGA25",
    "workflowId": "162",
    "pgmName": "Bachelor of Business Administration",
    "name": "Bachelor of Business Administration(with Credits)-Regular-FYUGP-2024-Business Administration"
  },
  "bba": {
    "code": "UCCAVHGS25",
    "workflowId": "96",
    "pgmName": "Bachelor of Business Administration",
    "name": "Bachelor of Business Administration(with Credits)-Regular-FYUGP-2024-Business Administration"
  },
  "aeg": {
    "code": "UCAENGGS25",
    "workflowId": "81",
    "pgmName": "Bachelor of Arts in English",
    "name": "Bachelor of Arts in English(with Credits)-Regular-FYUGP-2024-English"
  },
  "aec": {
    "code": "UCAECONS25",
    "workflowId": "81",
    "pgmName": "Bachelor of Arts in Economics",
    "name": "Bachelor of Arts in Social Science-Economics(with Credits)-Regular-FYUGP-2024-Social Science-Economics"
  },
  "bca": {
    "code": "CAP_BCA(HONOURS)",
    "workflowId": "122",
    "pgmName": "Bachelor of Computer Application",
    "name": "Bachelor of Computer Application(with Credits)-Regular-FYUGP-2024-Computer Application"
  },
  "ahi": {
    "code": "UCAHISGS25",
    "workflowId": "81",
    "pgmName": "Bachelor of Arts in History",
    "name": "Bachelor of Arts in History(with Credits)-Regular-FYUGP-2024-History"
  },
  "ttm": {
    "code": "UCTTMGS25",
    "workflowId": "96",
    "pgmName": "Bachelor of Tourism and Travel Management",
    "name": "Bachelor of Tourism and Travel Management(with Credits)-Regular-FYUGP-2024-TTM"
  },
  "ama": {
    "code": "UCAMALGS25",
    "workflowId": "81",
    "pgmName": "Bachelor of Arts in Malayalam",
    "name": "Bachelor of Arts in Malayalam(with Credits)-Regular-FYUGP-2024-Malayalam"
  },
  "ccs": {
    "code": "CSC_BSC(HONOURS)",
    "workflowId": "122",
    "pgmName": "Bachelor of Science in Computer Science",
    "name": "Bachelor of Science in Computer Science(with Credits)-Regular-FYUGP-2024-Computer Science"
  },
  "cps": {
    "code": "UCAPOLGS25",
    "workflowId": "81",
    "pgmName": "Bachelor of Arts in Political Science",
    "name": "Bachelor of Arts in Political Science(with Credits)-Regular-FYUGP-2024-Political Science"
  },
  "cms": {
    "code": "UCAMATHS25",
    "workflowId": "105",
    "pgmName": "Bachelor of Science in Mathematics",
    "name": "Bachelor of Science in Mathematics(with Credits)-Regular-FYUGP-2024-Mathematics"
  },
  "cph": {
    "code": "PHY_BSC(HONOURS)",
    "workflowId": "105",
    "pgmName": "Bachelor of Science in Physics",
    "name": "Bachelor of Science in Physics(with Credits)-Regular-FYUGP-2024-Physics"
  },
  "czo": {
    "code": "ZOO_BSC(HONOURS)",
    "workflowId": "105",
    "pgmName": "Bachelor of Science in Zoology",
    "name": "Bachelor of Science in Zoology(with Credits)-Regular-FYUGP-2024-Zoology"
  },
  "cch": {
    "code": "CHE_BSC(HONOURS)",
    "workflowId": "105",
    "pgmName": "Bachelor of Science in Chemistry",
    "name": "Bachelor of Science in Chemistry(with Credits)-Regular-FYUGP-2024-Chemistry"
  },
  "cbo": {
    "code": "BOT_BSC(HONOURS)",
    "workflowId": "105",
    "pgmName": "Bachelor of Science in Botany",
    "name": "Bachelor of Science in Botany(with Credits)-Regular-FYUGP-2024-Botany"
  },
  "akn": {
    "code": "KAN_BA(HONOURS)",
    "workflowId": "81",
    "pgmName": "Bachelor of Arts in Kannada",
    "name": "Bachelor of Arts in Kannada(with Credits)-Regular-FYUGP-2024-Kannada"
  },
  "bcom": {
    "code": "COM_BCOM(HONOURS)",
    "workflowId": "96",
    "pgmName": "Bachelor of Commerce",
    "name": "Bachelor of Commerce(with Credits)-Regular-FYUGP-2024-Commerce"
  }
};

export const KANNUR_COLLEGES = [
  {
    "id": "435",
    "code": "435",
    "name": "Department of Anthropology",
    "clgRef": "Department of Anthropology,Dr. Janaki Ammal Campus,Palayad"
  },
  {
    "id": "432",
    "code": "432",
    "name": "School of Behavioural Sciences, Mangattuparamba",
    "clgRef": "Department of Behavioural Sciences,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  {
    "id": "434",
    "code": "434",
    "name": "Department of Commerce and Business Studies",
    "clgRef": "Department of Commerce and Business Studies,DR.P.K.R.M.Campus, Kannur University,Palathadam P.O Nileswaram,Kasargod District,Pin:671314"
  },
  {
    "id": "430",
    "code": "430",
    "name": "Department of Information Technology, Mangattuparamba",
    "clgRef": "Jointly conducted by Dept of IT,Mathematics & Statistics,Mangattuparamba Campus, Kannur University Campus P. O, Kannur - 670 567"
  },
  {
    "id": "433",
    "code": "433",
    "name": "Department of Physics, Swami Anandatheertha Campus",
    "clgRef": "Jointly conducted by Dept of Physics & Chemistry,Swami Anandatheertha Campus, Payyannur P.O. Edat,Pin. 670 327"
  },
  {
    "id": "131",
    "code": "131",
    "name": "Department of Economics",
    "clgRef": "Department of Economics,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  {
    "id": "125",
    "code": "125",
    "name": "Department of Geography",
    "clgRef": "Department of Geography,Swami Anandatheertha Campus, Payyannur P.O. Edat, Pin. 670 327"
  },
  {
    "id": "133",
    "code": "133",
    "name": "Department of Management Studies",
    "clgRef": "Department of Management Studies,Thavakkara Campus,Thavakkara,pin:670002"
  },
  {
    "id": "128",
    "code": "128",
    "name": "Department of Environmental Studies",
    "clgRef": "Department of Environmental Science,Kannur University, Mangattuparamba Campus K U Campus P.O,Kannur,"
  },
  {
    "id": "124",
    "code": "124",
    "name": "Department of History",
    "clgRef": "Department of History,Kannur University, Thalassery Campus,Palayad P.O"
  },
  {
    "id": "130",
    "code": "130",
    "name": "Department of Studies in English",
    "clgRef": "Department of English,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670661"
  },
  {
    "id": "129",
    "code": "129",
    "name": "Department of Mathematical Sciences",
    "clgRef": "Jointly Conducted by Statistical Sciences and Mathematical Sciences,Mangattuparamba Campus,Kannur University Campus P. O, Kannur - 670 567"
  },
  {
    "id": "132",
    "code": "132",
    "name": "IT Education Centre",
    "clgRef": "Information Technology Education Center Palayad,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670"
  },
  {
    "id": "127",
    "code": "127",
    "name": "Department of Music",
    "clgRef": "Department of Music"
  },
  {
    "id": "431",
    "code": "431",
    "name": "School of Physical Education and Sport Sciences, Mangattuparamba",
    "clgRef": "School of Physical Education and  Sports Sciences"
  }
];

export const COLLEGE_LOOKUP = {
  "departmentofanthropologydrjanakiammalcampuspalayad": {
    "id": "435",
    "code": "435",
    "name": "Department of Anthropology",
    "clgRef": "Department of Anthropology,Dr. Janaki Ammal Campus,Palayad"
  },
  "department of anthropology,dr. janaki ammal campus,palayad": {
    "id": "435",
    "code": "435",
    "name": "Department of Anthropology",
    "clgRef": "Department of Anthropology,Dr. Janaki Ammal Campus,Palayad"
  },
  "departmentofanthropology": {
    "id": "435",
    "code": "435",
    "name": "Department of Anthropology",
    "clgRef": "Department of Anthropology,Dr. Janaki Ammal Campus,Palayad"
  },
  "department of anthropology": {
    "id": "435",
    "code": "435",
    "name": "Department of Anthropology",
    "clgRef": "Department of Anthropology,Dr. Janaki Ammal Campus,Palayad"
  },
  "435": {
    "id": "435",
    "code": "435",
    "name": "Department of Anthropology",
    "clgRef": "Department of Anthropology,Dr. Janaki Ammal Campus,Palayad"
  },
  "departmentofbehaviouralscienceskannuruniversitymangattuparambacampuskucampuspopin670567": {
    "id": "432",
    "code": "432",
    "name": "School of Behavioural Sciences, Mangattuparamba",
    "clgRef": "Department of Behavioural Sciences,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  "department of behavioural sciences,kannur university, mangattuparamba campus,k u campus p.o,pin : 670567": {
    "id": "432",
    "code": "432",
    "name": "School of Behavioural Sciences, Mangattuparamba",
    "clgRef": "Department of Behavioural Sciences,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  "schoolofbehaviouralsciencesmangattuparamba": {
    "id": "432",
    "code": "432",
    "name": "School of Behavioural Sciences, Mangattuparamba",
    "clgRef": "Department of Behavioural Sciences,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  "school of behavioural sciences, mangattuparamba": {
    "id": "432",
    "code": "432",
    "name": "School of Behavioural Sciences, Mangattuparamba",
    "clgRef": "Department of Behavioural Sciences,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  "432": {
    "id": "432",
    "code": "432",
    "name": "School of Behavioural Sciences, Mangattuparamba",
    "clgRef": "Department of Behavioural Sciences,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  "departmentofcommerceandbusinessstudiesdrpkrmcampuskannuruniversitypalathadamponileswaramkasargoddistrictpin671314": {
    "id": "434",
    "code": "434",
    "name": "Department of Commerce and Business Studies",
    "clgRef": "Department of Commerce and Business Studies,DR.P.K.R.M.Campus, Kannur University,Palathadam P.O Nileswaram,Kasargod District,Pin:671314"
  },
  "department of commerce and business studies,dr.p.k.r.m.campus, kannur university,palathadam p.o nileswaram,kasargod district,pin:671314": {
    "id": "434",
    "code": "434",
    "name": "Department of Commerce and Business Studies",
    "clgRef": "Department of Commerce and Business Studies,DR.P.K.R.M.Campus, Kannur University,Palathadam P.O Nileswaram,Kasargod District,Pin:671314"
  },
  "departmentofcommerceandbusinessstudies": {
    "id": "434",
    "code": "434",
    "name": "Department of Commerce and Business Studies",
    "clgRef": "Department of Commerce and Business Studies,DR.P.K.R.M.Campus, Kannur University,Palathadam P.O Nileswaram,Kasargod District,Pin:671314"
  },
  "department of commerce and business studies": {
    "id": "434",
    "code": "434",
    "name": "Department of Commerce and Business Studies",
    "clgRef": "Department of Commerce and Business Studies,DR.P.K.R.M.Campus, Kannur University,Palathadam P.O Nileswaram,Kasargod District,Pin:671314"
  },
  "434": {
    "id": "434",
    "code": "434",
    "name": "Department of Commerce and Business Studies",
    "clgRef": "Department of Commerce and Business Studies,DR.P.K.R.M.Campus, Kannur University,Palathadam P.O Nileswaram,Kasargod District,Pin:671314"
  },
  "jointlyconductedbydeptofitmathematicsstatisticsmangattuparambacampuskannuruniversitycampuspokannur670567": {
    "id": "430",
    "code": "430",
    "name": "Department of Information Technology, Mangattuparamba",
    "clgRef": "Jointly conducted by Dept of IT,Mathematics & Statistics,Mangattuparamba Campus, Kannur University Campus P. O, Kannur - 670 567"
  },
  "jointly conducted by dept of it,mathematics & statistics,mangattuparamba campus, kannur university campus p. o, kannur - 670 567": {
    "id": "430",
    "code": "430",
    "name": "Department of Information Technology, Mangattuparamba",
    "clgRef": "Jointly conducted by Dept of IT,Mathematics & Statistics,Mangattuparamba Campus, Kannur University Campus P. O, Kannur - 670 567"
  },
  "departmentofinformationtechnologymangattuparamba": {
    "id": "430",
    "code": "430",
    "name": "Department of Information Technology, Mangattuparamba",
    "clgRef": "Jointly conducted by Dept of IT,Mathematics & Statistics,Mangattuparamba Campus, Kannur University Campus P. O, Kannur - 670 567"
  },
  "department of information technology, mangattuparamba": {
    "id": "430",
    "code": "430",
    "name": "Department of Information Technology, Mangattuparamba",
    "clgRef": "Jointly conducted by Dept of IT,Mathematics & Statistics,Mangattuparamba Campus, Kannur University Campus P. O, Kannur - 670 567"
  },
  "430": {
    "id": "430",
    "code": "430",
    "name": "Department of Information Technology, Mangattuparamba",
    "clgRef": "Jointly conducted by Dept of IT,Mathematics & Statistics,Mangattuparamba Campus, Kannur University Campus P. O, Kannur - 670 567"
  },
  "jointlyconductedbydeptofphysicschemistryswamianandatheerthacampuspayyannurpoedatpin670327": {
    "id": "433",
    "code": "433",
    "name": "Department of Physics, Swami Anandatheertha Campus",
    "clgRef": "Jointly conducted by Dept of Physics & Chemistry,Swami Anandatheertha Campus, Payyannur P.O. Edat,Pin. 670 327"
  },
  "jointly conducted by dept of physics & chemistry,swami anandatheertha campus, payyannur p.o. edat,pin. 670 327": {
    "id": "433",
    "code": "433",
    "name": "Department of Physics, Swami Anandatheertha Campus",
    "clgRef": "Jointly conducted by Dept of Physics & Chemistry,Swami Anandatheertha Campus, Payyannur P.O. Edat,Pin. 670 327"
  },
  "departmentofphysicsswamianandatheerthacampus": {
    "id": "433",
    "code": "433",
    "name": "Department of Physics, Swami Anandatheertha Campus",
    "clgRef": "Jointly conducted by Dept of Physics & Chemistry,Swami Anandatheertha Campus, Payyannur P.O. Edat,Pin. 670 327"
  },
  "department of physics, swami anandatheertha campus": {
    "id": "433",
    "code": "433",
    "name": "Department of Physics, Swami Anandatheertha Campus",
    "clgRef": "Jointly conducted by Dept of Physics & Chemistry,Swami Anandatheertha Campus, Payyannur P.O. Edat,Pin. 670 327"
  },
  "433": {
    "id": "433",
    "code": "433",
    "name": "Department of Physics, Swami Anandatheertha Campus",
    "clgRef": "Jointly conducted by Dept of Physics & Chemistry,Swami Anandatheertha Campus, Payyannur P.O. Edat,Pin. 670 327"
  },
  "departmentofeconomicskannuruniversitymangattuparambacampuskucampuspopin670567": {
    "id": "131",
    "code": "131",
    "name": "Department of Economics",
    "clgRef": "Department of Economics,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  "department of economics,kannur university, mangattuparamba campus,k u campus p.o,pin : 670567": {
    "id": "131",
    "code": "131",
    "name": "Department of Economics",
    "clgRef": "Department of Economics,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  "departmentofeconomics": {
    "id": "131",
    "code": "131",
    "name": "Department of Economics",
    "clgRef": "Department of Economics,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  "department of economics": {
    "id": "131",
    "code": "131",
    "name": "Department of Economics",
    "clgRef": "Department of Economics,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  "131": {
    "id": "131",
    "code": "131",
    "name": "Department of Economics",
    "clgRef": "Department of Economics,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  "departmentofgeographyswamianandatheerthacampuspayyannurpoedatpin670327": {
    "id": "125",
    "code": "125",
    "name": "Department of Geography",
    "clgRef": "Department of Geography,Swami Anandatheertha Campus, Payyannur P.O. Edat, Pin. 670 327"
  },
  "department of geography,swami anandatheertha campus, payyannur p.o. edat, pin. 670 327": {
    "id": "125",
    "code": "125",
    "name": "Department of Geography",
    "clgRef": "Department of Geography,Swami Anandatheertha Campus, Payyannur P.O. Edat, Pin. 670 327"
  },
  "departmentofgeography": {
    "id": "125",
    "code": "125",
    "name": "Department of Geography",
    "clgRef": "Department of Geography,Swami Anandatheertha Campus, Payyannur P.O. Edat, Pin. 670 327"
  },
  "department of geography": {
    "id": "125",
    "code": "125",
    "name": "Department of Geography",
    "clgRef": "Department of Geography,Swami Anandatheertha Campus, Payyannur P.O. Edat, Pin. 670 327"
  },
  "125": {
    "id": "125",
    "code": "125",
    "name": "Department of Geography",
    "clgRef": "Department of Geography,Swami Anandatheertha Campus, Payyannur P.O. Edat, Pin. 670 327"
  },
  "departmentofmanagementstudiesthavakkaracampusthavakkarapin670002": {
    "id": "133",
    "code": "133",
    "name": "Department of Management Studies",
    "clgRef": "Department of Management Studies,Thavakkara Campus,Thavakkara,pin:670002"
  },
  "department of management studies,thavakkara campus,thavakkara,pin:670002": {
    "id": "133",
    "code": "133",
    "name": "Department of Management Studies",
    "clgRef": "Department of Management Studies,Thavakkara Campus,Thavakkara,pin:670002"
  },
  "departmentofmanagementstudies": {
    "id": "133",
    "code": "133",
    "name": "Department of Management Studies",
    "clgRef": "Department of Management Studies,Thavakkara Campus,Thavakkara,pin:670002"
  },
  "department of management studies": {
    "id": "133",
    "code": "133",
    "name": "Department of Management Studies",
    "clgRef": "Department of Management Studies,Thavakkara Campus,Thavakkara,pin:670002"
  },
  "133": {
    "id": "133",
    "code": "133",
    "name": "Department of Management Studies",
    "clgRef": "Department of Management Studies,Thavakkara Campus,Thavakkara,pin:670002"
  },
  "departmentofenvironmentalsciencekannuruniversitymangattuparambacampuskucampuspokannur": {
    "id": "128",
    "code": "128",
    "name": "Department of Environmental Studies",
    "clgRef": "Department of Environmental Science,Kannur University, Mangattuparamba Campus K U Campus P.O,Kannur,"
  },
  "department of environmental science,kannur university, mangattuparamba campus k u campus p.o,kannur,": {
    "id": "128",
    "code": "128",
    "name": "Department of Environmental Studies",
    "clgRef": "Department of Environmental Science,Kannur University, Mangattuparamba Campus K U Campus P.O,Kannur,"
  },
  "departmentofenvironmentalstudies": {
    "id": "128",
    "code": "128",
    "name": "Department of Environmental Studies",
    "clgRef": "Department of Environmental Science,Kannur University, Mangattuparamba Campus K U Campus P.O,Kannur,"
  },
  "department of environmental studies": {
    "id": "128",
    "code": "128",
    "name": "Department of Environmental Studies",
    "clgRef": "Department of Environmental Science,Kannur University, Mangattuparamba Campus K U Campus P.O,Kannur,"
  },
  "128": {
    "id": "128",
    "code": "128",
    "name": "Department of Environmental Studies",
    "clgRef": "Department of Environmental Science,Kannur University, Mangattuparamba Campus K U Campus P.O,Kannur,"
  },
  "departmentofhistorykannuruniversitythalasserycampuspalayadpo": {
    "id": "124",
    "code": "124",
    "name": "Department of History",
    "clgRef": "Department of History,Kannur University, Thalassery Campus,Palayad P.O"
  },
  "department of history,kannur university, thalassery campus,palayad p.o": {
    "id": "124",
    "code": "124",
    "name": "Department of History",
    "clgRef": "Department of History,Kannur University, Thalassery Campus,Palayad P.O"
  },
  "departmentofhistory": {
    "id": "124",
    "code": "124",
    "name": "Department of History",
    "clgRef": "Department of History,Kannur University, Thalassery Campus,Palayad P.O"
  },
  "department of history": {
    "id": "124",
    "code": "124",
    "name": "Department of History",
    "clgRef": "Department of History,Kannur University, Thalassery Campus,Palayad P.O"
  },
  "124": {
    "id": "124",
    "code": "124",
    "name": "Department of History",
    "clgRef": "Department of History,Kannur University, Thalassery Campus,Palayad P.O"
  },
  "departmentofenglishdrjanakiammalcampuspalayadthalasserypin670661": {
    "id": "130",
    "code": "130",
    "name": "Department of Studies in English",
    "clgRef": "Department of English,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670661"
  },
  "department of english,dr. janaki ammal campus,palayad,thalassery,pin : 670661": {
    "id": "130",
    "code": "130",
    "name": "Department of Studies in English",
    "clgRef": "Department of English,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670661"
  },
  "departmentofstudiesinenglish": {
    "id": "130",
    "code": "130",
    "name": "Department of Studies in English",
    "clgRef": "Department of English,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670661"
  },
  "department of studies in english": {
    "id": "130",
    "code": "130",
    "name": "Department of Studies in English",
    "clgRef": "Department of English,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670661"
  },
  "130": {
    "id": "130",
    "code": "130",
    "name": "Department of Studies in English",
    "clgRef": "Department of English,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670661"
  },
  "jointlyconductedbystatisticalsciencesandmathematicalsciencesmangattuparambacampuskannuruniversitycampuspokannur670567": {
    "id": "129",
    "code": "129",
    "name": "Department of Mathematical Sciences",
    "clgRef": "Jointly Conducted by Statistical Sciences and Mathematical Sciences,Mangattuparamba Campus,Kannur University Campus P. O, Kannur - 670 567"
  },
  "jointly conducted by statistical sciences and mathematical sciences,mangattuparamba campus,kannur university campus p. o, kannur - 670 567": {
    "id": "129",
    "code": "129",
    "name": "Department of Mathematical Sciences",
    "clgRef": "Jointly Conducted by Statistical Sciences and Mathematical Sciences,Mangattuparamba Campus,Kannur University Campus P. O, Kannur - 670 567"
  },
  "departmentofmathematicalsciences": {
    "id": "129",
    "code": "129",
    "name": "Department of Mathematical Sciences",
    "clgRef": "Jointly Conducted by Statistical Sciences and Mathematical Sciences,Mangattuparamba Campus,Kannur University Campus P. O, Kannur - 670 567"
  },
  "department of mathematical sciences": {
    "id": "129",
    "code": "129",
    "name": "Department of Mathematical Sciences",
    "clgRef": "Jointly Conducted by Statistical Sciences and Mathematical Sciences,Mangattuparamba Campus,Kannur University Campus P. O, Kannur - 670 567"
  },
  "129": {
    "id": "129",
    "code": "129",
    "name": "Department of Mathematical Sciences",
    "clgRef": "Jointly Conducted by Statistical Sciences and Mathematical Sciences,Mangattuparamba Campus,Kannur University Campus P. O, Kannur - 670 567"
  },
  "informationtechnologyeducationcenterpalayaddrjanakiammalcampuspalayadthalasserypin670": {
    "id": "132",
    "code": "132",
    "name": "IT Education Centre",
    "clgRef": "Information Technology Education Center Palayad,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670"
  },
  "information technology education center palayad,dr. janaki ammal campus,palayad,thalassery,pin : 670": {
    "id": "132",
    "code": "132",
    "name": "IT Education Centre",
    "clgRef": "Information Technology Education Center Palayad,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670"
  },
  "iteducationcentre": {
    "id": "132",
    "code": "132",
    "name": "IT Education Centre",
    "clgRef": "Information Technology Education Center Palayad,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670"
  },
  "it education centre": {
    "id": "132",
    "code": "132",
    "name": "IT Education Centre",
    "clgRef": "Information Technology Education Center Palayad,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670"
  },
  "132": {
    "id": "132",
    "code": "132",
    "name": "IT Education Centre",
    "clgRef": "Information Technology Education Center Palayad,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670"
  },
  "departmentofmusic": {
    "id": "127",
    "code": "127",
    "name": "Department of Music",
    "clgRef": "Department of Music"
  },
  "department of music": {
    "id": "127",
    "code": "127",
    "name": "Department of Music",
    "clgRef": "Department of Music"
  },
  "127": {
    "id": "127",
    "code": "127",
    "name": "Department of Music",
    "clgRef": "Department of Music"
  },
  "schoolofphysicaleducationandsportssciences": {
    "id": "431",
    "code": "431",
    "name": "School of Physical Education and Sport Sciences, Mangattuparamba",
    "clgRef": "School of Physical Education and  Sports Sciences"
  },
  "school of physical education and  sports sciences": {
    "id": "431",
    "code": "431",
    "name": "School of Physical Education and Sport Sciences, Mangattuparamba",
    "clgRef": "School of Physical Education and  Sports Sciences"
  },
  "schoolofphysicaleducationandsportsciencesmangattuparamba": {
    "id": "431",
    "code": "431",
    "name": "School of Physical Education and Sport Sciences, Mangattuparamba",
    "clgRef": "School of Physical Education and  Sports Sciences"
  },
  "school of physical education and sport sciences, mangattuparamba": {
    "id": "431",
    "code": "431",
    "name": "School of Physical Education and Sport Sciences, Mangattuparamba",
    "clgRef": "School of Physical Education and  Sports Sciences"
  },
  "431": {
    "id": "431",
    "code": "431",
    "name": "School of Physical Education and Sport Sciences, Mangattuparamba",
    "clgRef": "School of Physical Education and  Sports Sciences"
  },
  "313": {
    "id": "133",
    "code": "133",
    "name": "Department of Management Studies",
    "clgRef": "Department of Management Studies,Thavakkara Campus,Thavakkara,pin:670002"
  },
  "304": {
    "id": "434",
    "code": "434",
    "name": "Department of Commerce and Business Studies",
    "clgRef": "Department of Commerce and Business Studies,DR.P.K.R.M.Campus, Kannur University,Palathadam P.O Nileswaram,Kasargod District,Pin:671314"
  },
  "314": {
    "id": "128",
    "code": "128",
    "name": "Department of Environmental Studies",
    "clgRef": "Department of Environmental Science,Kannur University, Mangattuparamba Campus K U Campus P.O,Kannur,"
  },
  "312": {
    "id": "131",
    "code": "131",
    "name": "Department of Economics",
    "clgRef": "Department of Economics,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  "315": {
    "id": "124",
    "code": "124",
    "name": "Department of History",
    "clgRef": "Department of History,Kannur University, Thalassery Campus,Palayad P.O"
  },
  "316": {
    "id": "130",
    "code": "130",
    "name": "Department of Studies in English",
    "clgRef": "Department of English,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670661"
  },
  "317": {
    "id": "129",
    "code": "129",
    "name": "Department of Mathematical Sciences",
    "clgRef": "Jointly Conducted by Statistical Sciences and Mathematical Sciences,Mangattuparamba Campus,Kannur University Campus P. O, Kannur - 670 567"
  },
  "318": {
    "id": "132",
    "code": "132",
    "name": "IT Education Centre",
    "clgRef": "Information Technology Education Center Palayad,Dr. Janaki Ammal Campus,Palayad,Thalassery,Pin : 670"
  },
  "309": {
    "id": "125",
    "code": "125",
    "name": "Department of Geography",
    "clgRef": "Department of Geography,Swami Anandatheertha Campus, Payyannur P.O. Edat, Pin. 670 327"
  },
  "305": {
    "id": "432",
    "code": "432",
    "name": "School of Behavioural Sciences, Mangattuparamba",
    "clgRef": "Department of Behavioural Sciences,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567"
  },
  "306": {
    "id": "430",
    "code": "430",
    "name": "Department of Information Technology, Mangattuparamba",
    "clgRef": "Jointly conducted by Dept of IT,Mathematics & Statistics,Mangattuparamba Campus, Kannur University Campus P. O, Kannur - 670 567"
  },
  "307": {
    "id": "433",
    "code": "433",
    "name": "Department of Physics, Swami Anandatheertha Campus",
    "clgRef": "Jointly conducted by Dept of Physics & Chemistry,Swami Anandatheertha Campus, Payyannur P.O. Edat,Pin. 670 327"
  },
  "308": {
    "id": "435",
    "code": "435",
    "name": "Department of Anthropology",
    "clgRef": "Department of Anthropology,Dr. Janaki Ammal Campus,Palayad"
  },
  "323": {
    "id": "431",
    "code": "431",
    "name": "School of Physical Education and Sport Sciences, Mangattuparamba",
    "clgRef": "School of Physical Education and  Sports Sciences"
  },
  "200": {
    "id": "200",
    "code": "200",
    "name": "Govt. College Mananthavady",
    "clgRef": "Govt. College Mananthavady, Wayanad"
  },
  "govtcollegemananthavady": {
    "id": "200",
    "code": "200",
    "name": "Govt. College Mananthavady",
    "clgRef": "Govt. College Mananthavady, Wayanad"
  },
  "govtcollegemananthavadywayanad": {
    "id": "200",
    "code": "200",
    "name": "Govt. College Mananthavady",
    "clgRef": "Govt. College Mananthavady, Wayanad"
  },
  "201": {
    "id": "201",
    "code": "201",
    "name": "Govt. Brennen College, Thalassery",
    "clgRef": "Government Brennen College, Thalassery, Kannur"
  },
  "govtbrennencollegethalassery": {
    "id": "201",
    "code": "201",
    "name": "Govt. Brennen College, Thalassery",
    "clgRef": "Government Brennen College, Thalassery, Kannur"
  },
  "governmentbrennencollegethalasserykannur": {
    "id": "201",
    "code": "201",
    "name": "Govt. Brennen College, Thalassery",
    "clgRef": "Government Brennen College, Thalassery, Kannur"
  },
  "205": {
    "id": "205",
    "code": "205",
    "name": "Govt. College Kasaragod",
    "clgRef": "Govt College Kasargod, Kasargod"
  },
  "govtcollegekasaragod": {
    "id": "205",
    "code": "205",
    "name": "Govt. College Kasaragod",
    "clgRef": "Govt College Kasargod, Kasargod"
  },
  "govtcollegekasargodkasargod": {
    "id": "205",
    "code": "205",
    "name": "Govt. College Kasaragod",
    "clgRef": "Govt College Kasargod, Kasargod"
  },
  "207": {
    "id": "207",
    "code": "207",
    "name": "K.M.M. Govt. Women's College, Kannur",
    "clgRef": "K.M.M. Govt. Women's College, Kannur"
  },
  "kmmgovtwomenscollegekannur": {
    "id": "207",
    "code": "207",
    "name": "K.M.M. Govt. Women's College, Kannur",
    "clgRef": "K.M.M. Govt. Women's College, Kannur"
  },
  "209": {
    "id": "209",
    "code": "209",
    "name": "Payyannur College, Payyannur",
    "clgRef": "Payyannur College, Payyannur P, O., Edat, Kannur"
  },
  "payyannurcollegepayyannur": {
    "id": "209",
    "code": "209",
    "name": "Payyannur College, Payyannur",
    "clgRef": "Payyannur College, Payyannur P, O., Edat, Kannur"
  },
  "payyannurcollegepayyannurpoedatkannur": {
    "id": "209",
    "code": "209",
    "name": "Payyannur College, Payyannur",
    "clgRef": "Payyannur College, Payyannur P, O., Edat, Kannur"
  },
  "211": {
    "id": "211",
    "code": "211",
    "name": "Sir Syed College, Taliparamba",
    "clgRef": "Sir Syed College, Taliparamba, Kannur"
  },
  "sirsyedcollegetaliparamba": {
    "id": "211",
    "code": "211",
    "name": "Sir Syed College, Taliparamba",
    "clgRef": "Sir Syed College, Taliparamba, Kannur"
  },
  "sirsyedcollegetaliparambakannur": {
    "id": "211",
    "code": "211",
    "name": "Sir Syed College, Taliparamba",
    "clgRef": "Sir Syed College, Taliparamba, Kannur"
  },
  "212": {
    "id": "212",
    "code": "212",
    "name": "Sree Narayana College, Kannur",
    "clgRef": "Sree Narayana College, Kannur"
  },
  "sreenarayanacollegekannur": {
    "id": "212",
    "code": "212",
    "name": "Sree Narayana College, Kannur",
    "clgRef": "Sree Narayana College, Kannur"
  },
  "214": {
    "id": "214",
    "code": "214",
    "name": "Nirmalagiri College, Kuthuparamba",
    "clgRef": "Nirmalagiri College, Kuthuparamba, Kannur"
  },
  "nirmalagiricollegekuthuparamba": {
    "id": "214",
    "code": "214",
    "name": "Nirmalagiri College, Kuthuparamba",
    "clgRef": "Nirmalagiri College, Kuthuparamba, Kannur"
  },
  "nirmalagiricollegekuthuparambakannur": {
    "id": "214",
    "code": "214",
    "name": "Nirmalagiri College, Kuthuparamba",
    "clgRef": "Nirmalagiri College, Kuthuparamba, Kannur"
  },
  "215": {
    "id": "215",
    "code": "215",
    "name": "Mahatma Gandhi College, Iritty",
    "clgRef": "Mahatma Gandhi College, Iritty, Keezhur P.O., Kannur"
  },
  "mahatmagandhicollegeiritty": {
    "id": "215",
    "code": "215",
    "name": "Mahatma Gandhi College, Iritty",
    "clgRef": "Mahatma Gandhi College, Iritty, Keezhur P.O., Kannur"
  },
  "mahatmagandhicollegeirittykeezhurpokannur": {
    "id": "215",
    "code": "215",
    "name": "Mahatma Gandhi College, Iritty",
    "clgRef": "Mahatma Gandhi College, Iritty, Keezhur P.O., Kannur"
  },
  "216": {
    "id": "216",
    "code": "216",
    "name": "N.A.M. College, Kallikandy",
    "clgRef": "N.A.M. College, Kallikandy, Thuvakunnu P.O., Kannur"
  },
  "namcollegekallikandy": {
    "id": "216",
    "code": "216",
    "name": "N.A.M. College, Kallikandy",
    "clgRef": "N.A.M. College, Kallikandy, Thuvakunnu P.O., Kannur"
  },
  "namcollegekallikandythuvakunnupokannur": {
    "id": "216",
    "code": "216",
    "name": "N.A.M. College, Kallikandy",
    "clgRef": "N.A.M. College, Kallikandy, Thuvakunnu P.O., Kannur"
  },
  "241": {
    "id": "241",
    "code": "241",
    "name": "Nehru Arts and Science College, Kanhangad",
    "clgRef": "Nehru Arts and Science College, Kanhangad, Kasaragod"
  },
  "nehruartsandsciencecollegekanhangad": {
    "id": "241",
    "code": "241",
    "name": "Nehru Arts and Science College, Kanhangad",
    "clgRef": "Nehru Arts and Science College, Kanhangad, Kasaragod"
  },
  "nehruartsandsciencecollegekanhangadkasaragod": {
    "id": "241",
    "code": "241",
    "name": "Nehru Arts and Science College, Kanhangad",
    "clgRef": "Nehru Arts and Science College, Kanhangad, Kasaragod"
  },
  "251": {
    "id": "251",
    "code": "251",
    "name": "St. Pius X College, Rajapuram",
    "clgRef": "St. Pius X College, Rajapuram, Kasaragod"
  },
  "stpiusxcollegerajapuram": {
    "id": "251",
    "code": "251",
    "name": "St. Pius X College, Rajapuram",
    "clgRef": "St. Pius X College, Rajapuram, Kasaragod"
  },
  "stpiusxcollegerajapuramkasaragod": {
    "id": "251",
    "code": "251",
    "name": "St. Pius X College, Rajapuram",
    "clgRef": "St. Pius X College, Rajapuram, Kasaragod"
  },
  "265": {
    "id": "265",
    "code": "265",
    "name": "SES College, Sreekandapuram",
    "clgRef": "SES College, Sreekandapuram, Kannur"
  },
  "sescollegesreekandapuram": {
    "id": "265",
    "code": "265",
    "name": "SES College, Sreekandapuram",
    "clgRef": "SES College, Sreekandapuram, Kannur"
  },
  "sescollegesreekandapuramkannur": {
    "id": "265",
    "code": "265",
    "name": "SES College, Sreekandapuram",
    "clgRef": "SES College, Sreekandapuram, Kannur"
  },
  "266": {
    "id": "266",
    "code": "266",
    "name": "Wadihuda Institute of Research and Advanced Studies",
    "clgRef": "Wadihuda Institute of Research and Advanced Studies, Vilayancode"
  },
  "wadihudainstituteofresearchandadvancedstudies": {
    "id": "266",
    "code": "266",
    "name": "Wadihuda Institute of Research and Advanced Studies",
    "clgRef": "Wadihuda Institute of Research and Advanced Studies, Vilayancode"
  },
  "wadihudainstituteofresearchandadvancedstudiesvilayancode": {
    "id": "266",
    "code": "266",
    "name": "Wadihuda Institute of Research and Advanced Studies",
    "clgRef": "Wadihuda Institute of Research and Advanced Studies, Vilayancode"
  }
};

export const TARGET_COLUMNS = [
  'Name',
  'Gender',
  'Mobile',
  'Email',
  'username',
  'dob',
  'Date of Birth (DD/MM/YY)',
  'AadhaarNumber',
  'ABCId',
  'CountryId',
  'CountryName',
  'corrStateId',
  'corrStateName',
  'corrAddress',
  'corrCity',
  'corrPin',
  'categoryId',
  'categoryName',
  'CasteId',
  'Caste',
  'religionId',
  'religion',
  'PGM Name',
  'WORKFLOWID',
  'Program Name',
  'PROGRAMCODE',
  'CLG NAME',
  'CollegeId',
  'College Name',
  'CollegeCode',
  'QualificationFacultyId',
  'QualificationName',
  'QualificationSpecializationId',
  'Stream',
  'CertificateNumber',
  'NoOfAttempts',
  'MarksObtained',
  'Percentage',
  'MarksOutOf',
  'CGPA',
  'SeatNumber',
  '',
  'Result Date ',
  'ResultStatus'
];


export const LOOKUP_TAB_CONFIG = [
  {
    key: 'program',
    label: 'Programmes',
    type: 'program',
    nameHeader: 'Programme Name (Reference)',
    idHeader: 'ProgramCode / workFlowId',
    nameLabel: 'Programme Name (Reference)',
    idLabel: 'Program Details',
    description: 'Programme Name (Reference) ↔ ProgramCode ↔ workFlowId ↔ Program Name'
  },
  {
    key: 'college',
    label: 'Colleges',
    type: 'college',
    nameHeader: 'College Name',
    idHeader: 'CollegeId / CollegeCode',
    nameLabel: 'College Name',
    idLabel: 'College Details',
    description: 'CollegeId ↔ CollegeCode ↔ College Name ↔ CLG Name'
  },
  {
    key: 'category',
    label: 'Categories',
    type: 'simple',
    nameHeader: 'categoryName',
    idHeader: 'categoryId',
    nameLabel: 'Category Name',
    idLabel: 'Category ID',
    description: 'Social and reservation categories (e.g., SEBC-Muslim -> 10, Scheduled Castes -> 2, General -> 1)'
  },
  {
    key: 'caste',
    label: 'Castes',
    nameHeader: 'Caste',
    idHeader: 'CasteId',
    nameLabel: 'Caste',
    idLabel: 'Caste ID',
    description: 'Community & caste classifications (e.g., Mukkuvan -> 320, Thiyyas -> 169, Mappila -> 168)'
  },
  {
    key: 'religion',
    label: 'Religions',
    nameHeader: 'religion',
    idHeader: 'religionId',
    nameLabel: 'Religion',
    idLabel: 'Religion ID',
    description: 'Religion codes (e.g., Hindu -> 1, Islam -> 2, Christian -> 3, Others -> 5)'
  },
  {
    key: 'country',
    label: 'Countries',
    nameHeader: 'CountryName',
    idHeader: 'CountryId',
    nameLabel: 'Country Name',
    idLabel: 'Country ID',
    description: 'Nationality & country of origin (e.g., Indian / India -> 107, Nepal -> 108)'
  },
  {
    key: 'state',
    label: 'States',
    nameHeader: 'corrStateName',
    idHeader: 'corrStateId',
    nameLabel: 'State Name',
    idLabel: 'State ID',
    description: 'State of correspondence (e.g., Kerala -> 18, Karnataka -> 17, Tamil Nadu -> 31)'
  },
  {
    key: 'qualification',
    label: 'Qualifications',
    nameHeader: 'QualificationName',
    idHeader: 'QualificationFacultyId',
    nameLabel: 'Qualification Name',
    idLabel: 'Faculty ID',
    description: 'Qualifying boards & examinations (e.g., HSE - Kerala -> 2, CBSE -> 2, Others -> 7)'
  },
  {
    key: 'stream',
    label: 'Streams',
    nameHeader: 'Stream',
    idHeader: 'QualificationSpecializationId',
    nameLabel: 'Stream Name',
    idLabel: 'Specialization ID',
    description: 'Academic streams / groups (e.g., Science -> 1, Commerce -> 2, Humanities -> 3)'
  }
];

export const DEFAULT_LOOKUP_DATA = {
  program: KANNUR_PROGRAMMES,
  college: KANNUR_COLLEGES,
  category: [
    { name: 'SEBC-Other Backward Hindu', id: '10' },
    { name: 'SEBC-Muslim', id: '10' },
    { name: 'SEBC-Ezhava/Thiyya/Billava', id: '10' },
    { name: 'Scheduled Castes', id: '2' },
    { name: 'General', id: '1' },
    { name: 'EWS', id: '1' },
    { name: 'Scheduled Tribes', id: '3' },
    { name: 'SEBC-Latin Catholic(LC)', id: '10' }
  ],
  caste: [
  {
    "name": "Adi Karnataka",
    "id": "223"
  },
  {
    "name": "Adiyan",
    "id": "328"
  },
  {
    "name": "Agasa",
    "id": "2"
  },
  {
    "name": "Andhuru Nair",
    "id": "183"
  },
  {
    "name": "Arayan",
    "id": "6"
  },
  {
    "name": "Arunthathiyar",
    "id": "225"
  },
  {
    "name": "Arya",
    "id": "10"
  },
  {
    "name": "Asari",
    "id": "11"
  },
  {
    "name": "Baira",
    "id": "227"
  },
  {
    "name": "Bakuda",
    "id": "228"
  },
  {
    "name": "Bestha",
    "id": "15"
  },
  {
    "name": "Bhandari or Bhondari",
    "id": "16"
  },
  {
    "name": "Bharathar(Other than Parathar)",
    "id": "230"
  },
  {
    "name": "Billava",
    "id": "312"
  },
  {
    "name": "Bovies",
    "id": "17"
  },
  {
    "name": "Chakkiliyan",
    "id": "231"
  },
  {
    "name": "Chaliya",
    "id": "22"
  },
  {
    "name": "Cheramar",
    "id": "238"
  },
  {
    "name": "Cheruman",
    "id": "239"
  },
  {
    "name": "Chetty/Chetties",
    "id": "313"
  },
  {
    "name": "Converts from Scheduled Castes to Christianity",
    "id": "314"
  },
  {
    "name": "Devadiga",
    "id": "26"
  },
  {
    "name": "Devanga",
    "id": "27"
  },
  {
    "name": "Dheevara",
    "id": "28"
  },
  {
    "name": "Eruman",
    "id": "31"
  },
  {
    "name": "Ezhava",
    "id": "169"
  },
  {
    "name": "Ezhavas",
    "id": "316"
  },
  {
    "name": "Ezhuthachan",
    "id": "33"
  },
  {
    "name": "Ganaka",
    "id": "35"
  },
  {
    "name": "Gatti",
    "id": "38"
  },
  {
    "name": "General",
    "id": "393"
  },
  {
    "name": "Gowda",
    "id": "40"
  },
  {
    "name": "Hegde",
    "id": "42"
  },
  {
    "name": "Hindu Nadar",
    "id": "43"
  },
  {
    "name": "Irivavan",
    "id": "192"
  },
  {
    "name": "Irular",
    "id": "334"
  },
  {
    "name": "Islam",
    "id": "323"
  },
  {
    "name": "Jogi",
    "id": "50"
  },
  {
    "name": "Kadar (Wayanad Kadar)",
    "id": "336"
  },
  {
    "name": "Kalladi",
    "id": "247"
  },
  {
    "name": "Kammalas",
    "id": "59"
  },
  {
    "name": "Kanakkan",
    "id": "248"
  },
  {
    "name": "Kanisan",
    "id": "63"
  },
  {
    "name": "Kaniyan",
    "id": "65"
  },
  {
    "name": "Kannian",
    "id": "69"
  },
  {
    "name": "Karimpalan",
    "id": "339"
  },
  {
    "name": "Karuvan",
    "id": "72"
  },
  {
    "name": "Kattunayakan",
    "id": "340"
  },
  {
    "name": "Kavara",
    "id": "249"
  },
  {
    "name": "Kavuthiyan",
    "id": "74"
  },
  {
    "name": "Kitaran",
    "id": "79"
  },
  {
    "name": "Kollan",
    "id": "81"
  },
  {
    "name": "Koraga",
    "id": "343"
  },
  {
    "name": "Koteyar",
    "id": "320"
  },
  {
    "name": "Kottar Chetties",
    "id": "321"
  },
  {
    "name": "Kudiya",
    "id": "344"
  },
  {
    "name": "Kudumbi",
    "id": "85"
  },
  {
    "name": "Kulala",
    "id": "86"
  },
  {
    "name": "Kulalan",
    "id": "196"
  },
  {
    "name": "Kurava",
    "id": "254"
  },
  {
    "name": "Kuravan",
    "id": "255"
  },
  {
    "name": "Kurichchan (Kurichiyan)",
    "id": "347"
  },
  {
    "name": "Kuruman",
    "id": "348"
  },
  {
    "name": "Kurumans",
    "id": "349"
  },
  {
    "name": "Kusavan",
    "id": "90"
  },
  {
    "name": "Latin Catholics",
    "id": "322"
  },
  {
    "name": "Maila",
    "id": "257"
  },
  {
    "name": "Mala Vettuvan",
    "id": "358"
  },
  {
    "name": "Malayan",
    "id": "205"
  },
  {
    "name": "Malayarayar",
    "id": "365"
  },
  {
    "name": "Maniyani",
    "id": "95"
  },
  {
    "name": "Mannan",
    "id": "259"
  },
  {
    "name": "Mappila",
    "id": "323"
  },
  {
    "name": "Maratis of HosdurgTaluk",
    "id": "96"
  },
  {
    "name": "Maratis of Kasaragod and hosdurg Taluk",
    "id": "367"
  },
  {
    "name": "Maravan (Maravar)",
    "id": "97"
  },
  {
    "name": "Maruthuvar",
    "id": "98"
  },
  {
    "name": "Mavilan",
    "id": "368"
  },
  {
    "name": "Melakudi",
    "id": "369"
  },
  {
    "name": "Moger (Other than Mogeyar)",
    "id": "261"
  },
  {
    "name": "Mokaya",
    "id": "208"
  },
  {
    "name": "Moosari",
    "id": "108"
  },
  {
    "name": "Mukaya",
    "id": "109"
  },
  {
    "name": "Mukhari",
    "id": "110"
  },
  {
    "name": "Mukkuvan",
    "id": "111"
  },
  {
    "name": "Mundala",
    "id": "263"
  },
  {
    "name": "Muslim",
    "id": "168"
  },
  {
    "name": "Muslims",
    "id": "323"
  },
  {
    "name": "Muvari",
    "id": "112"
  },
  {
    "name": "Nadar belonging to Christian religious denominations other than SIUC",
    "id": "324"
  },
  {
    "name": "Naicken",
    "id": "113"
  },
  {
    "name": "Nair",
    "id": "1"
  },
  {
    "name": "Nalkadaya",
    "id": "265"
  },
  {
    "name": "Not Mentioned",
    "id": "393"
  },
  {
    "name": "Other",
    "id": "268"
  },
  {
    "name": "Padmasali",
    "id": "120"
  },
  {
    "name": "Paniyan",
    "id": "383"
  },
  {
    "name": "Paravan",
    "id": "277"
  },
  {
    "name": "Paraya",
    "id": "278"
  },
  {
    "name": "Parayan",
    "id": "279"
  },
  {
    "name": "Pathiyan",
    "id": "281"
  },
  {
    "name": "Perumannan",
    "id": "282"
  },
  {
    "name": "Perumkollen",
    "id": "129"
  },
  {
    "name": "Pulaya",
    "id": "284"
  },
  {
    "name": "Pulayan",
    "id": "286"
  },
  {
    "name": "Pulluvan",
    "id": "288"
  },
  {
    "name": "Rajapuri",
    "id": "131"
  },
  {
    "name": "Sadhu Chetty",
    "id": "132"
  },
  {
    "name": "Saliyas",
    "id": "136"
  },
  {
    "name": "Sambava",
    "id": "293"
  },
  {
    "name": "Scheduled Tribes (Lakshadweep)",
    "id": "393"
  },
  {
    "name": "Sidhanar",
    "id": "298"
  },
  {
    "name": "SIUC",
    "id": "326"
  },
  {
    "name": "Thandan",
    "id": "301"
  },
  {
    "name": "Thattan",
    "id": "144"
  },
  {
    "name": "Thiyya",
    "id": "169"
  },
  {
    "name": "Thiyyas",
    "id": "327"
  },
  {
    "name": "Thiyyas (who are known as Thandan, in the erstwhile Cochin and Malabar areas)",
    "id": "327"
  },
  {
    "name": "Uraly",
    "id": "391"
  },
  {
    "name": "Vakkaliga",
    "id": "155"
  },
  {
    "name": "Valan",
    "id": "157"
  },
  {
    "name": "Valluvan",
    "id": "304"
  },
  {
    "name": "Vanika Vaisya",
    "id": "160"
  },
  {
    "name": "Vaniya",
    "id": "161"
  },
  {
    "name": "Vaniyan",
    "id": "162"
  },
  {
    "name": "Vannan",
    "id": "305"
  },
  {
    "name": "Vannathan",
    "id": "163"
  },
  {
    "name": "Veera Saivas",
    "id": "164"
  },
  {
    "name": "Velan",
    "id": "306"
  },
  {
    "name": "Veluthedan",
    "id": "166"
  },
  {
    "name": "Veluthedathu Nair",
    "id": "167"
  },
  {
    "name": "Vetta Kuruman",
    "id": "392"
  },
  {
    "name": "Vettuvan",
    "id": "308"
  },
  {
    "name": "Vilakkathalavan",
    "id": "169"
  },
  {
    "name": "Vilakkithala Nair",
    "id": "170"
  },
  {
    "name": "Viswakarmas",
    "id": "175"
  },
  {
    "name": "WynadanChetty",
    "id": "176"
  },
  {
    "name": "Yadava",
    "id": "177"
  },
  {
    "name": "Yogi",
    "id": "180"
  }
],
  religion: [
    { name: 'Hindu', id: '1' },
    { name: 'Islam', id: '2' },
    { name: 'Christian', id: '3' },
    { name: 'Others', id: '5' }
  ],
  country: [
    { name: 'Indian', id: '107' },
    { name: 'India', id: '107' },
    { name: 'Nepal', id: '108' },
    { name: 'Bhutan', id: '109' },
    { name: 'Others', id: '999' }
  ],
  state: [
    { name: 'Kerala', id: '18' },
    { name: 'Karnataka', id: '17' },
    { name: 'Tamil Nadu', id: '31' },
    { name: 'Maharashtra', id: '20' },
    { name: 'Lakshadweep', id: '19' },
    { name: 'Andhra Pradesh', id: '2' },
    { name: 'Telangana', id: '32' },
    { name: 'Delhi', id: '10' },
    { name: 'Manipur', id: '21' },
    { name: 'Goa', id: '11' },
    { name: 'Gujarat', id: '12' },
    { name: 'Haryana', id: '13' },
    { name: 'Himachal Pradesh', id: '14' },
    { name: 'Jammu & Kashmir', id: '15' },
    { name: 'Jharkhand', id: '16' },
    { name: 'Madhya Pradesh', id: '19' },
    { name: 'Meghalaya', id: '22' },
    { name: 'Mizoram', id: '23' },
    { name: 'Nagaland', id: '24' },
    { name: 'Odisha', id: '25' },
    { name: 'Punjab', id: '26' },
    { name: 'Rajasthan', id: '27' },
    { name: 'Sikkim', id: '28' },
    { name: 'Tripura', id: '29' },
    { name: 'Uttar Pradesh', id: '33' },
    { name: 'Uttarakhand', id: '34' },
    { name: 'West Bengal', id: '35' },
    { name: 'Assam', id: '4' },
    { name: 'Bihar', id: '5' },
    { name: 'Chhattisgarh', id: '7' },
    { name: 'Chandigarh', id: '6' },
    { name: 'Puducherry', id: '25' },
    { name: 'Andaman and Nicobar Islands', id: '1' },
    { name: 'Ladakh', id: '37' }
  ],
  qualification: [
    { name: 'HSE - Kerala', id: '2' },
    { name: 'CBSE', id: '2' },
    { name: 'VHSE - Kerala/HSE NSQF', id: '2' },
    { name: 'ISC', id: '2' },
    { name: 'Others', id: '7' }
  ],
  stream: [
    { name: 'Commerce', id: '2' },
    { name: 'Humanities', id: '3' },
    { name: 'Science', id: '1' }
  ]
};

export function normalizeLookupKey(str) {
  return String(str || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

export function buildLookupMaps(lookupData) {
  const maps = {
    programByName: {},
    programByCode: {},
    collegeByName: {},
    collegeByCode: {},
    collegeById: {},
    categoryByName: {},
    categoryById: {},
    casteByName: {},
    casteById: {},
    religionByName: {},
    religionById: {},
    countryByName: {},
    countryById: {},
    stateByName: {},
    stateById: {},
    qualificationByName: {},
    qualificationById: {},
    streamByName: {},
    streamById: {}
  };

  const populateSimple = (list, byName, byId) => {
    if (!Array.isArray(list)) return;
    for (const item of list) {
      if (!item) continue;
      const name = String(item.name || '').trim();
      const id = String(item.id !== undefined && item.id !== null ? item.id : '').trim();
      if (name) {
        byName[name.toLowerCase()] = { id, name };
        byName[normalizeLookupKey(name)] = { id, name };
      }
      if (id) {
        byId[id] = { id, name };
      }
    }
  };

  const data = lookupData || DEFAULT_LOOKUP_DATA;

  // Populate Programmes
  if (Array.isArray(data.program)) {
    for (const p of data.program) {
      if (!p) continue;
      const pgmName = String(p.pgmName || '').trim();
      const code = String(p.code || '').trim();
      const workflowId = String(p.workflowId !== undefined && p.workflowId !== null ? p.workflowId : '').trim();
      const name = String(p.name || '').trim();
      const entry = { pgmName, code, workflowId, name };
      if (pgmName) {
        maps.programByName[pgmName.toLowerCase()] = entry;
        maps.programByName[normalizeLookupKey(pgmName)] = entry;
      }
      if (name) {
        maps.programByName[name.toLowerCase()] = entry;
        maps.programByName[normalizeLookupKey(name)] = entry;
      }
      if (code) {
        maps.programByCode[code.toLowerCase()] = entry;
        maps.programByCode[code] = entry;
      }
    }
  }

  // Populate Colleges
  if (Array.isArray(data.college)) {
    for (const c of data.college) {
      if (!c) continue;
      const id = String(c.id !== undefined && c.id !== null ? c.id : '').trim();
      const code = String(c.code !== undefined && c.code !== null ? c.code : '').trim();
      const name = String(c.name || '').trim();
      const clgRef = String(c.clgRef || '').trim();
      const entry = { id, code, name, clgRef };
      if (name) {
        maps.collegeByName[name.toLowerCase()] = entry;
        maps.collegeByName[normalizeLookupKey(name)] = entry;
      }
      if (clgRef) {
        maps.collegeByName[clgRef.toLowerCase()] = entry;
        maps.collegeByName[normalizeLookupKey(clgRef)] = entry;
      }
      if (code) {
        maps.collegeByCode[code] = entry;
        maps.collegeByCode[code.toLowerCase()] = entry;
      }
      if (id) {
        maps.collegeById[id] = entry;
      }
    }
  }

  // Populate Simple Lookups
  populateSimple(data.category, maps.categoryByName, maps.categoryById);
  populateSimple(data.caste, maps.casteByName, maps.casteById);
  populateSimple(data.religion, maps.religionByName, maps.religionById);
  populateSimple(data.country, maps.countryByName, maps.countryById);
  populateSimple(data.state, maps.stateByName, maps.stateById);
  populateSimple(data.qualification, maps.qualificationByName, maps.qualificationById);
  populateSimple(data.stream, maps.streamByName, maps.streamById);

  return maps;
}

const LOOKUP_STORAGE_KEY = 'report_studio_custom_lookups';

export function loadSavedLookups() {
  if (typeof window === 'undefined') return DEFAULT_LOOKUP_DATA;
  try {
    const raw = localStorage.getItem(LOOKUP_STORAGE_KEY);
    if (!raw) return DEFAULT_LOOKUP_DATA;
    const parsed = JSON.parse(raw);
    return {
      program: Array.isArray(parsed.program) && parsed.program.length > 0 ? parsed.program : DEFAULT_LOOKUP_DATA.program,
      college: Array.isArray(parsed.college) && parsed.college.length > 0 ? parsed.college : DEFAULT_LOOKUP_DATA.college,
      category: Array.isArray(parsed.category) && parsed.category.length > 0 ? parsed.category : DEFAULT_LOOKUP_DATA.category,
      caste: Array.isArray(parsed.caste) && parsed.caste.length > 0 ? parsed.caste : DEFAULT_LOOKUP_DATA.caste,
      religion: Array.isArray(parsed.religion) && parsed.religion.length > 0 ? parsed.religion : DEFAULT_LOOKUP_DATA.religion,
      country: Array.isArray(parsed.country) && parsed.country.length > 0 ? parsed.country : DEFAULT_LOOKUP_DATA.country,
      state: Array.isArray(parsed.state) && parsed.state.length > 0 ? parsed.state : DEFAULT_LOOKUP_DATA.state,
      qualification: Array.isArray(parsed.qualification) && parsed.qualification.length > 0 ? parsed.qualification : DEFAULT_LOOKUP_DATA.qualification,
      stream: Array.isArray(parsed.stream) && parsed.stream.length > 0 ? parsed.stream : DEFAULT_LOOKUP_DATA.stream
    };
  } catch (err) {
    console.warn('Failed to parse saved lookups from localStorage:', err);
    return DEFAULT_LOOKUP_DATA;
  }
}

export function saveLookupsToStorage(lookupData) {
  if (typeof window === 'undefined') return false;
  try {
    localStorage.setItem(LOOKUP_STORAGE_KEY, JSON.stringify(lookupData));
    return true;
  } catch (err) {
    console.error('Failed to save lookups to localStorage:', err);
    return false;
  }
}

export function parseUploadedLookupSpreadsheet(workbook, activeTabKey = 'caste') {
  if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error('Workbook is empty or contains no valid sheets.');
  }

  const pairsDef = [
    { key: 'category', nameKeys: ['categoryname', 'category', 'community'], idKeys: ['categoryid', 'catid'] },
    { key: 'caste', nameKeys: ['caste', 'subcaste', 'castename'], idKeys: ['casteid'] },
    { key: 'religion', nameKeys: ['religion', 'religionname'], idKeys: ['religionid', 'relid'] },
    { key: 'country', nameKeys: ['countryname', 'country', 'nationality'], idKeys: ['countryid'] },
    { key: 'state', nameKeys: ['corrstatename', 'statename', 'state', 'corrstate'], idKeys: ['corrstateid', 'stateid'] },
    { key: 'qualification', nameKeys: ['qualificationname', 'qualification', 'board', 'exam', 'qualifyingexam'], idKeys: ['qualificationfacultyid', 'facultyid', 'qualid'] },
    { key: 'stream', nameKeys: ['stream', 'streamname', 'specialization', 'branch'], idKeys: ['qualificationspecializationid', 'streamid', 'stream_id'] }
  ];

  const imported = {
    program: [],
    college: [],
    category: [],
    caste: [],
    religion: [],
    country: [],
    state: [],
    qualification: [],
    stream: []
  };

  let matchedSpecificPairs = 0;

  // Scan sheets
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!aoa || aoa.length < 2) continue;

    // Search for header row in first 5 rows
    for (let rIdx = 0; rIdx < Math.min(5, aoa.length); rIdx++) {
      const row = aoa[rIdx];
      if (!Array.isArray(row)) continue;
      const colMap = {};
      row.forEach((cell, idx) => {
        const norm = normalizeLookupKey(cell);
        if (norm) colMap[norm] = idx;
      });

      // Check Program columns
      const progRefCol = ['programmenamereference', 'pgmname', 'programreference', 'programmereference'].map(k => colMap[k]).find(idx => idx !== undefined);
      const progCodeCol = ['programcode', 'progcode', 'coursecode'].map(k => colMap[k]).find(idx => idx !== undefined);
      const progWfCol = ['workflowid', 'workflow'].map(k => colMap[k]).find(idx => idx !== undefined);
      const progNameCol = ['programname', 'fullname', 'name'].map(k => colMap[k]).find(idx => idx !== undefined && idx !== progRefCol);

      if ((progRefCol !== undefined || progNameCol !== undefined) && (progCodeCol !== undefined || progWfCol !== undefined)) {
        matchedSpecificPairs++;
        const seen = new Set(imported.program.map(x => x.pgmName.toLowerCase()));
        for (let r = rIdx + 1; r < aoa.length; r++) {
          const dataRow = aoa[r];
          if (!Array.isArray(dataRow)) continue;
          const pgmName = String(progRefCol !== undefined ? dataRow[progRefCol] : (dataRow[progNameCol] || '')).trim();
          const code = String(progCodeCol !== undefined ? dataRow[progCodeCol] : '').trim();
          const workflowId = String(progWfCol !== undefined ? dataRow[progWfCol] : '').trim();
          const name = String(progNameCol !== undefined ? dataRow[progNameCol] : pgmName).trim();
          if (pgmName && !seen.has(pgmName.toLowerCase())) {
            seen.add(pgmName.toLowerCase());
            imported.program.push({ pgmName, code, workflowId, name });
          }
        }
      }

      // Check College columns
      const clgIdCol = ['collegeid', 'colid'].map(k => colMap[k]).find(idx => idx !== undefined);
      const clgCodeCol = ['collegecode', 'colcode', 'centrecode'].map(k => colMap[k]).find(idx => idx !== undefined);
      const clgNameCol = ['collegename', 'college', 'centername'].map(k => colMap[k]).find(idx => idx !== undefined);
      const clgRefCol = ['clgname', 'collegeref', 'collegeaddress', 'campus'].map(k => colMap[k]).find(idx => idx !== undefined);

      if ((clgIdCol !== undefined || clgCodeCol !== undefined) && (clgNameCol !== undefined || clgRefCol !== undefined)) {
        matchedSpecificPairs++;
        const seen = new Set(imported.college.map(x => (x.name + '_' + x.id).toLowerCase()));
        for (let r = rIdx + 1; r < aoa.length; r++) {
          const dataRow = aoa[r];
          if (!Array.isArray(dataRow)) continue;
          const id = String(clgIdCol !== undefined ? dataRow[clgIdCol] : (clgCodeCol !== undefined ? dataRow[clgCodeCol] : '')).trim();
          const code = String(clgCodeCol !== undefined ? dataRow[clgCodeCol] : id).trim();
          const name = String(clgNameCol !== undefined ? dataRow[clgNameCol] : (clgRefCol !== undefined ? dataRow[clgRefCol] : '')).trim();
          const clgRef = String(clgRefCol !== undefined ? dataRow[clgRefCol] : name).trim();
          const dedupeKey = (name + '_' + id).toLowerCase();
          if (name && !seen.has(dedupeKey)) {
            seen.add(dedupeKey);
            imported.college.push({ id, code, name, clgRef });
          }
        }
      }

      // Check standard 2-column pairs
      for (const pair of pairsDef) {
        const nameCol = pair.nameKeys.map(k => colMap[k]).find(idx => idx !== undefined);
        const idCol = pair.idKeys.map(k => colMap[k]).find(idx => idx !== undefined);

        if (nameCol !== undefined && idCol !== undefined) {
          matchedSpecificPairs++;
          const seen = new Set(imported[pair.key].map(x => x.name.toLowerCase()));
          for (let r = rIdx + 1; r < aoa.length; r++) {
            const dataRow = aoa[r];
            if (!Array.isArray(dataRow)) continue;
            const name = String(dataRow[nameCol] || '').trim();
            const id = String(dataRow[idCol] !== undefined && dataRow[idCol] !== null ? dataRow[idCol] : '').trim();
            if (name && id && !seen.has(name.toLowerCase())) {
              seen.add(name.toLowerCase());
              imported[pair.key].push({ name, id });
            }
          }
        }
      }
    }
  }

  // Fallback: If no specific pairs matched but sheet has 2 columns, map to activeTabKey (if not program/college)
  if (matchedSpecificPairs === 0) {
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
    if (aoa && aoa.length >= 2) {
      const headerRow = aoa[0];
      let nameIdx = 0;
      let idIdx = 1;

      // Detect if second column looks like ID
      const h0 = String(headerRow[0] || '').toLowerCase();
      const h1 = String(headerRow[1] || '').toLowerCase();
      if (h0.includes('id') && !h1.includes('id')) {
        idIdx = 0;
        nameIdx = 1;
      }

      const targetKey = (activeTabKey && imported[activeTabKey] && activeTabKey !== 'program' && activeTabKey !== 'college') ? activeTabKey : 'caste';
      const seen = new Set();
      for (let r = 1; r < aoa.length; r++) {
        const dataRow = aoa[r];
        if (!Array.isArray(dataRow)) continue;
        const name = String(dataRow[nameIdx] || '').trim();
        const id = String(dataRow[idIdx] !== undefined && dataRow[idIdx] !== null ? dataRow[idIdx] : '').trim();
        if (name && id && !seen.has(name.toLowerCase())) {
          seen.add(name.toLowerCase());
          imported[targetKey].push({ name, id });
        }
      }
    }
  }

  const summary = {};
  let total = 0;
  for (const k of Object.keys(imported)) {
    summary[k] = imported[k].length;
    total += imported[k].length;
  }

  return { imported, summary, total };
}
