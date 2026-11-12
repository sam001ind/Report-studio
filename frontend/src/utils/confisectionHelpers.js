export const CONFIDENTIAL_SECTION_COLUMNS = [
  'Event Name',
  'Center Code',
  'Center Name',
  'Venue Code',
  'Venue Name',
  'Program Name',
  'Exam Date',
  'Course  Start Time',
  'Course  End Time',
  'Course  Code',
  'Course  Name',
  'Student Count',
  'venueId'
];

// Helper to calculate question paper packets
export const calculateQpPackets = (studentCount, bufferRule = 'plus1') => {
  const count = parseInt(studentCount, 10) || 0;
  if (count <= 0) {
    return { p20: 0, p10: 0, p5: 0, singles: 0, buffer: 0, totalQp: 0 };
  }

  let buffer = 0;
  if (bufferRule === 'plus1') {
    buffer = 1;
  } else if (bufferRule === 'plus2') {
    buffer = 2;
  } else if (bufferRule === 'five_percent') {
    buffer = Math.max(1, Math.ceil(count * 0.05));
  } else if (bufferRule === 'none') {
    buffer = 0;
  }

  const totalQp = count + buffer;

  const p20 = Math.floor(totalQp / 20);
  let rem = totalQp % 20;

  const p10 = Math.floor(rem / 10);
  rem = rem % 10;

  const p5 = Math.floor(rem / 5);
  const singles = rem % 5;

  return {
    p20,
    p10,
    p5,
    singles,
    buffer,
    totalQp
  };
};

// Helper for Session detection
export const getExamSession = (startTime, endTime) => {
  const s = `${startTime || ''} ${endTime || ''}`.toUpperCase();
  if (s.includes('AM')) return 'FN';
  if (s.includes('PM')) {
    if (s.includes('09:') || s.includes('10:') || s.includes('11:')) return 'FN';
    return 'AN';
  }
  return 'FN';
};
