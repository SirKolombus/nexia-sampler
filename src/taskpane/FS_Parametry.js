// FS_Parametry — seznam všech kombinací pro druhou úlohu
// Každý záznam má tvar: { controlRisk, inherentRisk, analyticalTests, controlTests, factor }
// Pole "factor" je zde vyplněno textem "DOPLŇ" - doplň prosím požadované číslo pro každou kombinaci.

const rows = [
  { controlRisk: 'Nízké', inherentRisk: 'Nízké', rmmLevel: 'Nízké', analyticalTests: 'Ano', controlTests: 'Ano', factor: '0' },
  { controlRisk: 'Nízké', inherentRisk: 'Nízké', rmmLevel: 'Nízké', analyticalTests: 'Ano', controlTests: 'Ne', factor: 'Test kontrol Error' },
  { controlRisk: 'Nízké', inherentRisk: 'Nízké', rmmLevel: 'Nízké', analyticalTests: 'Ne', controlTests: 'Ano', factor: '0.7' },
  { controlRisk: 'Nízké', inherentRisk: 'Nízké', rmmLevel: 'Nízké', analyticalTests: 'Ne', controlTests: 'Ne', factor: 'Test kontrol Error' },

  { controlRisk: 'Nízké', inherentRisk: 'Střední', rmmLevel: 'Nízké', analyticalTests: 'Ano', controlTests: 'Ano', factor: '0' },
  { controlRisk: 'Nízké', inherentRisk: 'Střední', rmmLevel: 'Nízké', analyticalTests: 'Ano', controlTests: 'Ne', factor: 'Test kontrol Error' },
  { controlRisk: 'Nízké', inherentRisk: 'Střední', rmmLevel: 'Nízké', analyticalTests: 'Ne', controlTests: 'Ano', factor: '1.1' },
  { controlRisk: 'Nízké', inherentRisk: 'Střední', rmmLevel: 'Nízké', analyticalTests: 'Ne', controlTests: 'Ne', factor: 'Test kontrol Error' },

  { controlRisk: 'Nízké', inherentRisk: 'Vysoké', rmmLevel: 'Střední', analyticalTests: 'Ano', controlTests: 'Ano', factor: '0.7' },
  { controlRisk: 'Nízké', inherentRisk: 'Vysoké', rmmLevel: 'Střední', analyticalTests: 'Ano', controlTests: 'Ne', factor: 'Test kontrol Error' },
  { controlRisk: 'Nízké', inherentRisk: 'Vysoké', rmmLevel: 'Střední', analyticalTests: 'Ne', controlTests: 'Ano', factor: '1.6' },
  { controlRisk: 'Nízké', inherentRisk: 'Vysoké', rmmLevel: 'Střední', analyticalTests: 'Ne', controlTests: 'Ne', factor: 'Test kontrol Error' },

  { controlRisk: 'Střední', inherentRisk: 'Nízké', rmmLevel: 'Nízké', analyticalTests: 'Ano', controlTests: 'Ano', factor: '0.7' },
  { controlRisk: 'Střední', inherentRisk: 'Nízké', rmmLevel: 'Nízké', analyticalTests: 'Ano', controlTests: 'Ne', factor: 'Test kontrol Error' },
  { controlRisk: 'Střední', inherentRisk: 'Nízké', rmmLevel: 'Nízké', analyticalTests: 'Ne', controlTests: 'Ano', factor: '1.4' },
  { controlRisk: 'Střední', inherentRisk: 'Nízké', rmmLevel: 'Nízké', analyticalTests: 'Ne', controlTests: 'Ne', factor: 'Test kontrol Error' },

  { controlRisk: 'Střední', inherentRisk: 'Střední', rmmLevel: 'Střední', analyticalTests: 'Ano', controlTests: 'Ano', factor: '1.1' },
  { controlRisk: 'Střední', inherentRisk: 'Střední', rmmLevel: 'Střední', analyticalTests: 'Ano', controlTests: 'Ne', factor: 'Test kontrol Error' },
  { controlRisk: 'Střední', inherentRisk: 'Střední', rmmLevel: 'Střední', analyticalTests: 'Ne', controlTests: 'Ano', factor: '1.8' },
  { controlRisk: 'Střední', inherentRisk: 'Střední', rmmLevel: 'Střední', analyticalTests: 'Ne', controlTests: 'Ne', factor: 'Test kontrol Error' },

  { controlRisk: 'Střední', inherentRisk: 'Vysoké', rmmLevel: 'Střední', analyticalTests: 'Ano', controlTests: 'Ano', factor: '1.4' },
  { controlRisk: 'Střední', inherentRisk: 'Vysoké', rmmLevel: 'Střední', analyticalTests: 'Ano', controlTests: 'Ne', factor: 'Test kontrol Error' },
  { controlRisk: 'Střední', inherentRisk: 'Vysoké', rmmLevel: 'Střední', analyticalTests: 'Ne', controlTests: 'Ano', factor: '2.3' },
  { controlRisk: 'Střední', inherentRisk: 'Vysoké', rmmLevel: 'Střední', analyticalTests: 'Ne', controlTests: 'Ne', factor: 'Test kontrol Error' },

  { controlRisk: 'Vysoké', inherentRisk: 'Nízké', rmmLevel: 'Nízké', analyticalTests: 'Ano', controlTests: 'Ano', factor: 'Test kontrol Error' },
  { controlRisk: 'Vysoké', inherentRisk: 'Nízké', rmmLevel: 'Nízké', analyticalTests: 'Ano', controlTests: 'Ne', factor: '1.1' },
  { controlRisk: 'Vysoké', inherentRisk: 'Nízké', rmmLevel: 'Nízké', analyticalTests: 'Ne', controlTests: 'Ano', factor: 'Test kontrol Error' },
  { controlRisk: 'Vysoké', inherentRisk: 'Nízké', rmmLevel: 'Nízké', analyticalTests: 'Ne', controlTests: 'Ne', factor: '1.9' },

  { controlRisk: 'Vysoké', inherentRisk: 'Střední', rmmLevel: 'Střední', analyticalTests: 'Ano', controlTests: 'Ano', factor: 'Test kontrol Error' },
  { controlRisk: 'Vysoké', inherentRisk: 'Střední', rmmLevel: 'Střední', analyticalTests: 'Ano', controlTests: 'Ne', factor: '1.4' },
  { controlRisk: 'Vysoké', inherentRisk: 'Střední', rmmLevel: 'Střední', analyticalTests: 'Ne', controlTests: 'Ano', factor: 'Test kontrol Error' },
  { controlRisk: 'Vysoké', inherentRisk: 'Střední', rmmLevel: 'Střední', analyticalTests: 'Ne', controlTests: 'Ne', factor: '2.3' },

  { controlRisk: 'Vysoké', inherentRisk: 'Vysoké', rmmLevel: 'Vysoké', analyticalTests: 'Ano', controlTests: 'Ano', factor: 'Test kontrol Error' },
  { controlRisk: 'Vysoké', inherentRisk: 'Vysoké', rmmLevel: 'Vysoké', analyticalTests: 'Ano', controlTests: 'Ne', factor: '1.9' },
  { controlRisk: 'Vysoké', inherentRisk: 'Vysoké', rmmLevel: 'Vysoké', analyticalTests: 'Ne', controlTests: 'Ano', factor: 'Test kontrol Error' },
  { controlRisk: 'Vysoké', inherentRisk: 'Vysoké', rmmLevel: 'Vysoké', analyticalTests: 'Ne', controlTests: 'Ne', factor: '3' }
];

export default rows;
