import * as XLSX from 'xlsx';
import { MemorialPerson } from '../types';

export const fetchAndParseData = async (): Promise<MemorialPerson[]> => {
  try {
    // Attempt to fetch from public folder
    const response = await fetch('/victims_data.xlsx');
    if (!response.ok) {
        throw new Error('Local XLSX not found');
    }
    
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    
    // Assuming the first sheet contains the data
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    let data = XLSX.utils.sheet_to_json<MemorialPerson>(worksheet);

    // Add coordinates if not present
    data = data.map((p, i) => {
        // Pre-generate offset lat/lng for mapping
        const t = Math.random();
        const w = (Math.random() - 0.5) * 0.08;
        
        return {
            ...p,
            // Defaults if they are not in the XLSX
            Name: p.Name || (p as any)['الاسم باللغة الانجليزية'] || "Unknown",
            الاسم: p.الاسم || "غير معروف",
            Sex: p.Sex || (p as any)['الجنس'] || (i % 2 === 0 ? "m" : "f"),
            Age: p.Age || (p as any)['العمر'] || "20",
            ID: p.ID || (p as any)['رقم الهوية'] || "",
            lat: p.lat ?? (31.23 + (t * 0.34) + w * -0.66), 
            lng: p.lng ?? (34.22 + (t * 0.30) + w * 0.75)
        };
    });
    
    return data;
  } catch (error) {
     console.error('Failed to load XLSX:', error);
     // Fallback mock data
     return Array.from({length: 150}).map((_, i) => {
        const t = Math.random();
        const w = (Math.random() - 0.5) * 0.08;
        return {
          Index: String(i+1),
          Name: "Mock Data (Please add victims_data.xlsx to public folder)",
          الاسم: "بيانات تجريبية (أضف ملف victims_data.xlsx إلى مجلد public)",
          Age: i % 7 === 0 ? "10" : (i % 3 === 0 ? "65" : "25"),
          Sex: i % 2 === 0 ? "m" : "f",
          ID: "00000",
          lat: 31.23 + (t * 0.34) + w * -0.66,
          lng: 34.22 + (t * 0.30) + w * 0.75
        };
     });
  }
};
