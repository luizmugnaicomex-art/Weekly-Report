import React, { useState, useEffect, useMemo } from "react";
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import * as XLSX from 'xlsx';
import { BrokerRow, TruckSupplierRow, PaymentRow, PaymentCategory, VesselRow, LIRow, QuotaEntry, VesselDetail } from "./types";
import { statusPillClasses, paymentStatusClasses } from "./constants";
import { Card } from "./components/Card";
import { NumberInput } from "./components/NumberInput";
import { TextAreaField } from "./components/TextAreaField";

// Helper to convert Excel date values (serial or DD/MM/YYYY) to YYYY-MM-DD for <input type="date">
const formatExcelDate = (val: any): string => {
  if (!val) return "";
  if (val instanceof Date) return val.toISOString().split('T')[0];
  if (typeof val === 'number') {
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    return date.toISOString().split('T')[0];
  }
  if (typeof val === 'string' && val.includes('/')) {
    const parts = val.split('/');
    if (parts.length === 3) {
      let [d, m, y] = parts;
      if (y.length === 2) y = "20" + y;
      return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
    }
  }
  if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(val)) return val;
  return "";
};

const parseBrazilianNumber = (val: any): number => {
  if (val === undefined || val === null || val === "") return 0;
  if (typeof val === 'number') return val;
  let str = String(val).trim();
  if (str.includes('.') && str.includes(',')) str = str.replace(/\./g, '').replace(',', '.');
  else if (str.includes(',')) str = str.replace(',', '.');
  return parseFloat(str.replace(/[^0-9.-]+/g, "")) || 0;
};

const App: React.FC = () => {
  const projectName = "Weekly Report - SUPERVISAO";
  const [activeTab, setActiveTab] = useState<"weekly" | "logistics" | "broker" | "financial">("weekly");
  const [isExporting, setIsExporting] = useState(false);
  const [expandedVesselId, setExpandedVesselId] = useState<string | null>(null);

  // Global Report Metadata
  const [weekLabel, setWeekLabel] = useState("Week 41 – 2026-10-05 to 2026-10-10");
  const [reportDate, setReportDate] = useState("2026-10-05");
  const [author, setAuthor] = useState("Luiz Vieira");
  const [authorTitle, setAuthorTitle] = useState("Customs & Logistics Supervisor");
  const [department, setDepartment] = useState("Inbound & Trade Compliance");
  const [directorName, setDirectorName] = useState("Operations Director");

  // Logistics KPIs & State
  const [containersArrived, setContainersArrived] = useState(1420);
  const [avgLeadTime, setAvgLeadTime] = useState(5.4);
  const [truckAvailability, setTruckAvailability] = useState(96);
  const [buffer, setBuffer] = useState(120);

  const [evBalance, setEvBalance] = useState(34175461.08);
  const [phevBalance, setPhevBalance] = useState(106057725.65);

  const [lis, setLis] = useState<LIRow[]>([
    { id: "1", liNumber: "25/4606635-0", poNumber: "9900003751", model: "120 Dolphin Mini", status: "Exigência atendida" },
    { id: "2", liNumber: "25/4612390-2", poNumber: "9900003780", model: "BYD Song Plus", status: "Registered" }
  ]);

  const [nextVessels, setNextVessels] = useState<VesselRow[]>([
    { id: "1", vesselName: "GREEN ITAJAI", quantity: 33, eta: "2026-10-12", status: "CARGO READY", warehouse: "TECON", deliveryScheduleDate: "2026-10-15", deadlineDate: "2026-11-28", update: "Expected in clear weather.", details: [] },
    { id: "2", vesselName: "BYD SUPREME", quantity: 45, eta: "2026-10-18", status: "IN TRANSIT", warehouse: "PARANAGUA", deliveryScheduleDate: "2026-10-21", deadlineDate: "2026-12-05", update: "Vessel on schedule.", details: [] }
  ]);

  const [quotas, setQuotas] = useState<QuotaEntry[]>([
    { id: "q1", vessel: "GREEN TAICANG", po: "9900003650", ncm: "8703.80.00", type: "EV", round: 1, li: "25/4750587-0", registerDate: "2026-10-01", fobValue: 8884615.39, qty: 840, exigenciaDate: "2026-10-03", docsDate: "2026-10-03", status: "RELEASED", di: "25/2849832-1", channel: "GREEN" },
  ]);

  const [quotaFilters, setQuotaFilters] = useState<Partial<Record<keyof QuotaEntry, string>>>({});
  
  const filteredQuotas = useMemo(() => {
    return quotas.filter((q) => {
      return Object.entries(quotaFilters).every(([key, value]) => {
        if (!value) return true;
        const val = q[key as keyof QuotaEntry];
        return String(val ?? "").toLowerCase().includes(String(value).toLowerCase());
      });
    });
  }, [quotas, quotaFilters]);

  const evTotalQty = useMemo(() => filteredQuotas.filter(q => q.type === "EV").reduce((acc, cur) => acc + cur.qty, 0), [filteredQuotas]);
  const phevTotalQty = useMemo(() => filteredQuotas.filter(q => q.type === "PHEV").reduce((acc, cur) => acc + cur.qty, 0), [filteredQuotas]);
  const evTotalSpent = useMemo(() => filteredQuotas.filter(q => q.type === "EV").reduce((acc, cur) => acc + cur.fobValue, 0), [filteredQuotas]);
  const phevTotalSpent = useMemo(() => filteredQuotas.filter(q => q.type === "PHEV").reduce((acc, cur) => acc + cur.fobValue, 0), [filteredQuotas]);

  // Summaries & Reports for Weekly Report Tab
  const [executiveSummary, setExecutiveSummary] = useState("Stable operation across all inbound channels. Customs clearance times maintained within target SLA (under 6 days). Zero demurrage incurred.");
  const [highlights, setHighlights] = useState("- Cleared 96% of arriving containers on first pass.\n- Successfully processed Round 1 EV quota allocations.\n- Zero production stoppages at plant.");
  const [issues, setIssues] = useState("- Supplier X allocation scheduling delay for battery components.\n- Port congestion at terminal berth 4 during peak weekend.");
  const [nextActions, setNextActions] = useState("- Executive alignment meeting with Supplier X.\n- Coordinate priority berthing schedule with TECON port authority.");
  const [declarationsSummary, setDeclarationsSummary] = useState("Mainly auto parts and complete knock-down (CKD) vehicle kits.");
  
  // Tab 2: Logistics Department Situation Report for Director
  const [logisticsDirectorReport, setLogisticsDirectorReport] = useState(
    `TO: ${directorName}\nFROM: ${author} (${authorTitle})\nDATE: ${reportDate}\nSUBJECT: Weekly Logistics & Inbound Operations Status Report\n\n1. EXECUTIVE SUMMARY & VESSEL PIPELINE:\n- Total arrived containers for the week reached 1,420 units with an average lead time of 5.4 days.\n- Vessel 'GREEN ITAJAI' (33 containers) is scheduled to berth at TECON with clearance docs pre-cleared.\n- Container yard buffer is maintained at a healthy 120 units, preventing any plant line starvation.\n\n2. TRANSPORT & TRUCK FLEET HEALTH:\n- Transport supplier allocation rate stands at 96% fulfillment.\n- Zero equipment shortages reported across primary transport corridors.\n\n3. CURRENT BOTTLENECK & ACTION PLAN:\n- Minor berthing queue at Paranaguá handled via priority discharge agreements.`
  );

  // Tab 3: Broker Department Situation Report for Director
  const [brokerDirectorReport, setBrokerDirectorReport] = useState(
    `TO: ${directorName}\nFROM: ${author} (${authorTitle})\nDATE: ${reportDate}\nSUBJECT: Weekly Customs Broker & Clearance Status Report\n\n1. CLEARANCE PERFORMANCE SUMMARY:\n- Customs broker network operating efficiently with 74% Green Channel release rate across primary brokers.\n- Total active LI and DI processes monitored without critical delays.\n\n2. BROKER OPERATIONS & EXIGÊNCIAS:\n- Broker Alpha processed 58 declarations with 5.8 days average clearance.\n- All pending exigências were successfully resolved within the 24-hour response SLA.\n\n3. ESCALATIONS & NEXT STEPS:\n- Regular alignment scheduled with broker supervisors to streamline NCM classification audits.`
  );
  
  const [brokers, setBrokers] = useState<BrokerRow[]>([
    { id: "1", name: "Broker Alpha (Santos)", processes: 58, avgClearanceDays: 5.8, greenChannelRate: 74, issues: "Highly efficient performance.", status: "OK" },
    { id: "2", name: "Broker Beta (Paranaguá)", processes: 34, avgClearanceDays: 6.1, greenChannelRate: 68, issues: "Minor documentation backlog cleared.", status: "Attention" }
  ]);

  const [truckDeclarationsSummary, setTruckDeclarationsSummary] = useState("Fleet capacity stable with 96% allocation fulfillment.");
  const [suppliers, setSuppliers] = useState<TruckSupplierRow[]>([
    { id: "1", supplierName: "Transporter Alpha Logistics", allocationRate: 98, onTimeRate: 96, comments: "Best performance in container hauling." }
  ]);

  // Tab 4: Financial Department Situation Report for Director
  const [financialDirectorReport, setFinancialDirectorReport] = useState(
    `TO: ${directorName}\nFROM: ${author} (${authorTitle})\nDATE: ${reportDate}\nSUBJECT: Weekly Inbound Financial Control & Quota Budget Report\n\n1. QUOTA BALANCE & FOB SPEND:\n- EV Quota Remaining Balance: $${(evBalance - evTotalSpent).toLocaleString('en-US', { minimumFractionDigits: 2 })}\n- PHEV Quota Remaining Balance: $${(phevBalance - phevTotalSpent).toLocaleString('en-US', { minimumFractionDigits: 2 })}\n- Expenditure is strictly aligned with approved quarterly import limits.\n\n2. PAYMENT & INVOICING OBLIGATIONS:\n- Warehouse storage and demurrage fees are fully up to date with zero overdue invoices.\n- Scheduled truck carrier and broker payments are queued for weekly disbursement.\n\n3. FINANCIAL OUTLOOK:\n- Favorable exchange rate hedging secured for upcoming inventory batches.`
  );

  const [payments, setPayments] = useState<PaymentRow[]>([
    { id: "p1", category: "Warehouse", entity: "TECON Rio Grande", description: "Storage & handling batch #401", plannedAmount: 145000, pendingAmount: 0, currency: "BRL", status: "Paid" },
    { id: "p2", category: "Truck Carrier", entity: "Transporter Alpha", description: "Container haulage weekly billing", plannedAmount: 210000, pendingAmount: 45000, currency: "BRL", status: "Scheduled" },
    { id: "p3", category: "Broker", entity: "Broker Alpha", description: "Customs clearance fees", plannedAmount: 85000, pendingAmount: 0, currency: "BRL", status: "Paid" },
    { id: "p4", category: "Other Invoices", entity: "Port Authority", description: "Terminal handling surcharges", plannedAmount: 32000, pendingAmount: 12000, currency: "USD", status: "Scheduled" }
  ]);
  const paymentCategories: PaymentCategory[] = ["Warehouse", "Truck Carrier", "Broker", "Other Invoices"];

  const STORAGE_KEY = "LOGISTICS_REPORT_DATA_QUOTA_V16";

  const handleSave = () => {
    const data = { 
      weekLabel, reportDate, author, authorTitle, department, directorName, containersArrived, avgLeadTime, truckAvailability, buffer, 
      lis, nextVessels, quotas, executiveSummary, highlights, issues, nextActions, declarationsSummary, 
      logisticsDirectorReport, brokerDirectorReport, financialDirectorReport,
      brokers, truckDeclarationsSummary, suppliers, payments, evBalance, phevBalance 
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    alert("All report data saved locally successfully.");
  };

  useEffect(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const p = JSON.parse(saved);
        if (p.weekLabel) setWeekLabel(p.weekLabel);
        if (p.reportDate) setReportDate(p.reportDate);
        if (p.author) setAuthor(p.author);
        if (p.authorTitle) setAuthorTitle(p.authorTitle);
        if (p.department) setDepartment(p.department);
        if (p.directorName) setDirectorName(p.directorName);
        if (p.lis) setLis(p.lis);
        if (p.nextVessels) setNextVessels(p.nextVessels);
        if (p.quotas) setQuotas(p.quotas);
        if (p.brokers) setBrokers(p.brokers);
        if (p.logisticsDirectorReport) setLogisticsDirectorReport(p.logisticsDirectorReport);
        if (p.brokerDirectorReport) setBrokerDirectorReport(p.brokerDirectorReport);
        if (p.financialDirectorReport) setFinancialDirectorReport(p.financialDirectorReport);
        if (p.suppliers) setSuppliers(p.suppliers);
        if (p.payments) setPayments(p.payments);
        if (p.executiveSummary) setExecutiveSummary(p.executiveSummary);
        if (p.highlights) setHighlights(p.highlights);
        if (p.issues) setIssues(p.issues);
        if (p.nextActions) setNextActions(p.nextActions);
        if (p.containersArrived !== undefined) setContainersArrived(p.containersArrived);
        if (p.avgLeadTime !== undefined) setAvgLeadTime(p.avgLeadTime);
        if (p.truckAvailability !== undefined) setTruckAvailability(p.truckAvailability);
        if (p.buffer !== undefined) setBuffer(p.buffer);
        if (p.evBalance) setEvBalance(p.evBalance);
        if (p.phevBalance) setPhevBalance(p.phevBalance);
      } catch (e) { console.error(e); }
    }
  }, []);

  const generatePDF = async () => {
    const element = document.getElementById('report-content');
    if (!element) return;
    setIsExporting(true);
    try {
        const clone = element.cloneNode(true) as HTMLElement;
        clone.style.position = 'absolute'; clone.style.left = '-9999px'; clone.style.width = '1250px'; clone.style.padding = '40px'; clone.style.backgroundColor = '#ffffff';
        document.body.appendChild(clone);
        const originalInputs = element.querySelectorAll('input, textarea, select');
        const cloneInputs = clone.querySelectorAll('input, textarea, select');
        cloneInputs.forEach((el, index) => {
            const original = originalInputs[index] as HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement;
            const parent = el.parentNode;
            if (!parent) return;
            let replacement: HTMLElement;
            if (el.tagName === 'SELECT') {
                replacement = document.createElement('span');
                replacement.textContent = (original as HTMLSelectElement).options[(original as HTMLSelectElement).selectedIndex]?.text || original.value;
                replacement.style.cssText = 'font-size: 9px; font-weight: bold; border: 1px solid #cbd5e1; padding: 1px 4px; border-radius: 4px; background: #f8fafc;';
            } else if (el.tagName === 'TEXTAREA') {
                replacement = document.createElement('div');
                replacement.textContent = original.value;
                replacement.style.cssText = 'white-space: pre-wrap; font-size: 11px; padding: 6px; line-height: 1.5;';
            } else {
                replacement = document.createElement('span');
                replacement.textContent = original.value;
                replacement.style.fontSize = '10px';
            }
            parent.replaceChild(replacement, el);
        });
        clone.querySelectorAll('button, input[type="file"], .filter-input').forEach(el => el.remove());
        const canvas = await html2canvas(clone, { scale: 2, useCORS: true, windowWidth: 1250 });
        document.body.removeChild(clone);
        const imgData = canvas.toDataURL('image/png', 1.0);
        const pdf = new jsPDF('p', 'mm', 'a4');
        const pdfWidth = pdf.internal.pageSize.getWidth();
        const imgHeight = (canvas.height * pdfWidth) / canvas.width;
        pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, imgHeight);
        pdf.save(`Weekly_Report_SUPERVISAO_${activeTab.toUpperCase()}_Export.pdf`);
    } catch (e) { console.error(e); } finally { setIsExporting(false); }
  };

  const handleQuotaFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const targetSheetNames = ["Quotas", "QUOTAS", "LI", "Li"];
      let wsName = wb.SheetNames[0];
      for (const name of targetSheetNames) {
        if (wb.SheetNames.includes(name)) {
          wsName = name;
          break;
        }
      }
      const ws = wb.Sheets[wsName];
      const data = XLSX.utils.sheet_to_json(ws);
      const mapped = data.map((row: any) => {
        const roundStr = String(row['ROUND'] || row['Round'] || "1");
        const roundNum = parseInt(roundStr.replace(/[^0-9]/g, "")) || 1;
        return {
          id: Math.random().toString(36).substr(2, 9),
          vessel: String(row['VESSEL'] || row['Vessel'] || "").trim(),
          po: String(row['PO'] || ""),
          ncm: String(row['NCM'] || ""),
          type: String(row['TYPE'] || row['Type'] || "EV").trim().toUpperCase(),
          round: roundNum,
          li: String(row['LI'] || ""),
          registerDate: formatExcelDate(row['REGISTER DATE'] || row['Register Date']),
          fobValue: parseBrazilianNumber(row['FOB VALUE (USD)'] || row['FOB VALUE'] || row['VALUE']),
          qty: parseInt(String(row['QTY CAR'] || row['QTY'] || 0)),
          exigenciaDate: formatExcelDate(row['EXIGENCIA D'] || row['EXIGENCIA DATE']),
          docsDate: formatExcelDate(row['DOCS DATE']),
          status: String(row['STATUS'] || "").trim(),
          di: String(row['DI'] || ""),
          channel: String(row['CHANNEL'] || "").trim().toUpperCase(),
        };
      });
      setQuotas(mapped);
    };
    reader.readAsBinaryString(file);
  };

  const handleVesselFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const targetSheetNames = ["Next Vessels", "NEXT VESSELS", "Vessels", "VESSELS", "Shipments"];
      let wsName = wb.SheetNames[0];
      for (const name of targetSheetNames) {
        if (wb.SheetNames.includes(name)) {
          wsName = name;
          break;
        }
      }
      const ws = wb.Sheets[wsName];
      const data = XLSX.utils.sheet_to_json(ws);
      const groups: Record<string, VesselRow> = {};

      data.forEach((row: any) => {
        const vesselNameRaw = String(row['Vessel'] || row['ARRIVAL VESSEL'] || "UNKNOWN").trim();
        const vesselKey = vesselNameRaw.toUpperCase();
        const qtyVal = row['BL Qty'] || row['CONTAINER QTY'] || row['Qty'] || 0;
        const qty = parseInt(String(qtyVal).replace(/[^0-9]+/g, "")) || 0;
        
        const detail: VesselDetail = {
          id: Math.random().toString(36).substr(2, 9),
          blNumber: String(row['BL'] || row['SAP BL/AWB'] || "N/A").trim(),
          po: String(row['PO Number'] || row['PO SAP'] || "N/A").trim(),
          poTotalBatch: String(row['PO Total C Batch'] || "").trim(),
          voyage: String(row['VOYAGE'] || "N/A").trim(),
          qty: qty,
          cargoType: String(row['Cargo Type'] || row['SHIPMENT TYPE OF CARGO'] || "FCL").trim(),
          status: String(row['Status'] || row['STATUS'] || "CARGO READY").trim(),
          broker: String(row['Broker'] || "").trim(),
        };

        if (!groups[vesselKey]) {
          groups[vesselKey] = {
            id: Math.random().toString(36).substr(2, 9),
            vesselName: vesselNameRaw,
            quantity: 0,
            eta: formatExcelDate(row['ETA'] || row['ACTUAL ETA']),
            status: String(row['Status'] || row['STATUS'] || "").trim(),
            warehouse: String(row['Warehouse'] || row['BONDED WAREHOUSE'] || "").trim(),
            deadlineDate: formatExcelDate(row['FREE TIME DEADLINE'] || ""),
            deliveryScheduleDate: "",
            update: "",
            details: [],
          };
        }
        groups[vesselKey].quantity += qty;
        groups[vesselKey].details?.push(detail);
      });
      
      setNextVessels(Object.values(groups).map(v => ({ ...v, update: `Aggregated from ${v.details?.length} BLs.` })));
    };
    reader.readAsBinaryString(file);
  };

  const toggleVesselExpand = (id: string) => setExpandedVesselId(expandedVesselId === id ? null : id);
  const handleQuotaChange = (id: string, field: keyof QuotaEntry, value: string | number) => setQuotas(prev => prev.map(q => q.id === id ? { ...q, [field]: value } : q));
  const handleAddQuota = () => setQuotas([...quotas, { id: Math.random().toString(36).substr(2, 9), vessel: "", po: "", ncm: "", type: "EV", round: 1, li: "", registerDate: "", fobValue: 0, qty: 0, exigenciaDate: "", docsDate: "", status: "", di: "", channel: "" }]);
  const handleRemoveQuota = (id: string) => setQuotas(quotas.filter(q => q.id !== id));

  const calculateRoundTotals = (type: string, round: number) => {
    const items = filteredQuotas.filter(q => q.type === type && q.round === round);
    const qty = items.reduce((acc, cur) => acc + cur.qty, 0);
    const value = items.reduce((acc, cur) => acc + cur.fobValue, 0);
    return { qty, value, items };
  };

  const sortedVessels = useMemo(() => [...nextVessels].sort((a, b) => (a.eta || '9999').localeCompare(b.eta || '9999')), [nextVessels]);

  const handleVesselChange = (id: string, field: keyof VesselRow, value: string | number) => setNextVessels(prev => prev.map(v => v.id === id ? { ...v, [field]: value } : v));
  
  const handleVesselDetailChange = (vesselId: string, detailId: string, field: keyof VesselDetail, value: string | number) => {
    setNextVessels(prev => prev.map(v => {
      if (v.id === vesselId && v.details) {
        const updatedDetails = v.details.map(d => d.id === detailId ? { ...d, [field]: value } : d);
        const totalQty = updatedDetails.reduce((acc, cur) => acc + (Number(cur.qty) || 0), 0);
        return { ...v, details: updatedDetails, quantity: totalQty };
      }
      return v;
    }));
  };

  const handleAddVesselDetail = (vesselId: string) => {
    setNextVessels(prev => prev.map(v => {
      if (v.id === vesselId) {
        const newDetail: VesselDetail = { id: Math.random().toString(36).substr(2, 9), blNumber: "N/A", po: "N/A", poTotalBatch: "", voyage: "N/A", qty: 0, cargoType: "FCL", status: "CARGO READY", broker: "" };
        return { ...v, details: [...(v.details || []), newDetail] };
      }
      return v;
    }));
  };

  const handleRemoveVesselDetail = (vesselId: string, detailId: string) => {
    setNextVessels(prev => prev.map(v => {
      if (v.id === vesselId && v.details) {
        const updatedDetails = v.details.filter(d => d.id !== detailId);
        const totalQty = updatedDetails.reduce((acc, cur) => acc + (Number(cur.qty) || 0), 0);
        return { ...v, details: updatedDetails, quantity: totalQty };
      }
      return v;
    }));
  };

  const handleAddVessel = () => setNextVessels([...nextVessels, { id: Math.random().toString(36).substr(2, 9), vesselName: "", quantity: 0, eta: "", status: "", warehouse: "", deliveryScheduleDate: "", deadlineDate: "", update: "", details: [] }]);
  const handleRemoveVessel = (id: string) => setNextVessels(nextVessels.filter(v => v.id !== id));
  const handleBrokerChange = (id: string, field: keyof BrokerRow, value: string | number) => setBrokers(prev => prev.map(b => b.id === id ? { ...b, [field]: value } : b));
  const handleSupplierChange = (id: string, field: keyof TruckSupplierRow, value: string | number) => setSuppliers(prev => prev.map(s => s.id === id ? { ...s, [field]: value } : s));
  const handlePaymentChange = (id: string, field: keyof PaymentRow, value: string | number) => setPayments(prev => prev.map(p => p.id === id ? { ...p, [field]: value } : p));
  const handleAddPayment = (category: PaymentCategory) => setPayments([...payments, { id: Math.random().toString(36).substr(2, 9), category, entity: "", description: "", plannedAmount: 0, pendingAmount: 0, currency: "BRL", status: "Scheduled" }]);

  const RoundTable = ({ type, round }: { type: string, round: number }) => {
    const { qty, value, items } = calculateRoundTotals(type, round);
    if (items.length === 0) return null;
    return (
      <div className="mb-6">
        <div className="bg-yellow-400 font-black text-[10px] px-3 py-1 uppercase border-x border-t border-slate-300 rounded-t-lg shadow-sm">{round} ROUND</div>
        <div className="overflow-x-auto border-x border-b border-slate-300 rounded-b-lg">
          <table className="w-full text-[9px] border-collapse min-w-[600px]">
            <thead className="bg-yellow-400/80">
              <tr className="uppercase font-bold text-slate-900 border-b border-slate-400">
                <th className="p-1 border-r border-slate-300 text-left w-[18%]">Vessel</th>
                <th className="p-1 border-r border-slate-300 text-center w-[12%]">Reg Date</th>
                <th className="p-1 border-r border-slate-300 text-right w-[8%]">Qty</th>
                <th className="p-1 border-r border-slate-300 text-right w-[15%]">Value (USD)</th>
                <th className="p-1 border-r border-slate-300 text-center w-[10%]">LI</th>
                <th className="p-1 border-r border-slate-300 text-center w-[10%]">DI</th>
                <th className="p-1 border-r border-slate-300 text-center w-[8%]">CH</th>
                <th className="p-1 border-r border-slate-300 text-left w-[15%]">Status</th>
                <th className="p-1 w-6 no-print"></th>
              </tr>
            </thead>
            <tbody className="bg-white">
              {items.map(q => (
                <tr key={q.id} className="hover:bg-slate-50 transition-colors group">
                  <td className="p-1 border-r border-slate-200"><input className="w-full bg-transparent outline-none font-bold" value={q.vessel} onChange={e => handleQuotaChange(q.id, 'vessel', e.target.value)} /></td>
                  <td className="p-1 border-r border-slate-200 text-center"><input type="date" className="w-full bg-transparent outline-none text-center" value={q.registerDate} onChange={e => handleQuotaChange(q.id, 'registerDate', e.target.value)} /></td>
                  <td className="p-1 border-r border-slate-200 text-right"><input type="number" className="w-full bg-transparent outline-none text-right font-bold" value={q.qty} onChange={e => handleQuotaChange(q.id, 'qty', Number(e.target.value))} /></td>
                  <td className="p-1 border-r border-slate-200 text-right font-mono"><input type="number" className="w-full bg-transparent outline-none text-right" value={q.fobValue} onChange={e => handleQuotaChange(q.id, 'fobValue', Number(e.target.value))} /></td>
                  <td className="p-1 border-r border-slate-200 text-center"><input className="w-full bg-transparent outline-none text-center font-medium" value={q.li} onChange={e => handleQuotaChange(q.id, 'li', e.target.value)} /></td>
                  <td className="p-1 border-r border-slate-200 text-center"><input className="w-full bg-transparent outline-none text-center font-medium" value={q.di} onChange={e => handleQuotaChange(q.id, 'di', e.target.value)} /></td>
                  <td className="p-1 border-r border-slate-200 text-center">
                     <select className={`w-full bg-transparent outline-none text-center font-black ${q.channel === 'GREEN' ? 'text-emerald-600' : q.channel === 'YELLOW' ? 'text-amber-600' : q.channel === 'RED' ? 'text-rose-600' : ''}`} value={q.channel} onChange={e => handleQuotaChange(q.id, 'channel', e.target.value)}>
                        <option value="">-</option>
                        <option value="GREEN">G</option>
                        <option value="YELLOW">Y</option>
                        <option value="RED">R</option>
                     </select>
                  </td>
                  <td className="p-1 border-r border-slate-200"><input className="w-full bg-transparent outline-none italic text-[8px]" value={q.status} onChange={e => handleQuotaChange(q.id, 'status', e.target.value)} /></td>
                  <td className="text-center no-print"><button onClick={() => handleRemoveQuota(q.id)} className="text-slate-300 hover:text-red-500"><span className="material-icons text-xs">close</span></button></td>
                </tr>
              ))}
            </tbody>
            <tfoot className="bg-yellow-300/40 font-black">
              <tr>
                <td colSpan={2} className="p-1 text-center uppercase">Total Round {round}</td>
                <td className="p-1 text-right text-sm">{qty}</td>
                <td className="p-1 text-right text-sm">${value.toLocaleString('en-US', { minimumFractionDigits: 2 })}</td>
                <td className="p-1" colSpan={5}>-</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-[#f1f5f9] text-slate-800 font-sans pb-16">
      {/* Top Header & 4 Tabs */}
      <header className="border-b border-slate-200 bg-white sticky top-0 z-50 no-print shadow-sm">
        <div className="max-w-[1450px] mx-auto px-6 py-3 flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-3 w-full md:w-auto justify-between md:justify-start">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl bg-red-600 flex items-center justify-center text-white font-black text-lg shadow-sm">B</div>
              <div className="flex flex-col">
                <span className="text-[10px] uppercase tracking-widest text-slate-500 font-bold leading-none mb-1">Weekly Supervision</span>
                <span className="font-extrabold text-slate-900 tracking-tight text-sm">Weekly Report - SUPERVISAO</span>
              </div>
            </div>
          </div>

          {/* 4 Exact Tabs */}
          <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-xl border border-slate-200 shadow-inner overflow-x-auto max-w-full">
            <button 
              onClick={() => setActiveTab("weekly")} 
              className={`px-4 py-2 text-xs font-bold rounded-lg transition-all flex items-center gap-2 whitespace-nowrap ${activeTab === 'weekly' ? 'bg-white text-slate-900 shadow-sm border border-slate-200/60' : 'text-slate-600 hover:text-slate-900'}`}
            >
              <span className="material-icons text-sm text-red-600">assignment</span>
              Weekly Report (Entire)
            </button>
            <button 
              onClick={() => setActiveTab("logistics")} 
              className={`px-4 py-2 text-xs font-bold rounded-lg transition-all flex items-center gap-2 whitespace-nowrap ${activeTab === 'logistics' ? 'bg-white text-slate-900 shadow-sm border border-slate-200/60' : 'text-slate-600 hover:text-slate-900'}`}
            >
              <span className="material-icons text-sm text-sky-600">local_shipping</span>
              Logistics
            </button>
            <button 
              onClick={() => setActiveTab("broker")} 
              className={`px-4 py-2 text-xs font-bold rounded-lg transition-all flex items-center gap-2 whitespace-nowrap ${activeTab === 'broker' ? 'bg-white text-slate-900 shadow-sm border border-slate-200/60' : 'text-slate-600 hover:text-slate-900'}`}
            >
              <span className="material-icons text-sm text-blue-600">admin_panel_settings</span>
              Broker
            </button>
            <button 
              onClick={() => setActiveTab("financial")} 
              className={`px-4 py-2 text-xs font-bold rounded-lg transition-all flex items-center gap-2 whitespace-nowrap ${activeTab === 'financial' ? 'bg-white text-slate-900 shadow-sm border border-slate-200/60' : 'text-slate-600 hover:text-slate-900'}`}
            >
              <span className="material-icons text-sm text-emerald-600">payments</span>
              Financial
            </button>
          </div>

          <div className="flex items-center gap-3 w-full md:w-auto justify-end">
            <button onClick={handleSave} className="px-4 py-2 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded-xl text-xs font-bold hover:bg-emerald-100 shadow-sm transition-all flex items-center gap-1.5">
              <span className="material-icons text-sm">save</span>Save Data
            </button>
            <button onClick={generatePDF} disabled={isExporting} className="px-4 py-2 bg-slate-900 text-white rounded-xl text-xs font-bold shadow-md hover:bg-slate-800 disabled:opacity-50 flex items-center gap-2 transition-all">
              <span className="material-icons text-sm">{isExporting ? 'sync' : 'picture_as_pdf'}</span>
              {isExporting ? 'Exporting...' : 'Export PDF'}
            </button>
          </div>
        </div>
      </header>

      {/* Main Report Container */}
      <main id="report-content" className="max-w-[1450px] mx-auto px-6 py-8 flex flex-col gap-6">
        
        {/* Professional Header & Metadata */}
        <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
          <div className="flex flex-col gap-1.5">
            <div className="flex items-center gap-2">
              <span className="px-2.5 py-0.5 bg-red-50 text-red-700 font-bold text-[10px] rounded uppercase tracking-wider border border-red-200">Director Briefing Suite</span>
              <span className="text-xs text-slate-400 font-medium">· {department}</span>
            </div>
            <h1 className="text-xl md:text-2xl font-black text-slate-900 tracking-tight">
              {activeTab === 'weekly' && "Entire Weekly Operations & Customs Supervision Report"}
              {activeTab === 'logistics' && "Logistics Department Situation Report for Director"}
              {activeTab === 'broker' && "Broker Department Situation Report for Director"}
              {activeTab === 'financial' && "Financial Department Situation Report for Director"}
            </h1>
            <p className="text-xs text-slate-500 font-medium">
              {activeTab === 'weekly' && "Comprehensive data tables, KPIs, vessel pipelines, and operational summaries."}
              {activeTab === 'logistics' && "Full written situation report regarding inbound logistics, vessels, and transport performance."}
              {activeTab === 'broker' && "Full written situation report regarding customs brokerage, clearance times, and channel rates."}
              {activeTab === 'financial' && "Full written situation report regarding budget balances, FOB spend, and payment obligations."}
            </p>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 bg-slate-50 p-3 rounded-xl border border-slate-200 w-full md:w-auto">
            <label className="flex flex-col gap-0.5 text-[10px] font-bold text-slate-500 uppercase">
              Week Scope
              <input type="text" value={weekLabel} onChange={e => setWeekLabel(e.target.value)} className="bg-white border border-slate-200 rounded px-2 py-1 text-xs font-bold text-slate-900 outline-none" />
            </label>
            <label className="flex flex-col gap-0.5 text-[10px] font-bold text-slate-500 uppercase">
              Report Date
              <input type="date" value={reportDate} onChange={e => setReportDate(e.target.value)} className="bg-white border border-slate-200 rounded px-2 py-1 text-xs font-medium text-slate-900 outline-none text-center" />
            </label>
            <label className="flex flex-col gap-0.5 text-[10px] font-bold text-slate-500 uppercase">
              Author
              <input type="text" value={author} onChange={e => setAuthor(e.target.value)} className="bg-white border border-slate-200 rounded px-2 py-1 text-xs font-bold text-slate-900 outline-none" />
            </label>
            <label className="flex flex-col gap-0.5 text-[10px] font-bold text-slate-500 uppercase">
              Director Recipient
              <input type="text" value={directorName} onChange={e => setDirectorName(e.target.value)} className="bg-white border border-slate-200 rounded px-2 py-1 text-xs font-bold text-slate-900 outline-none" />
            </label>
          </div>
        </div>

        {/* ========================================================= */}
        {/* TAB 1: WEEKLY REPORT ENTIRE (Summary of all + Full Data)  */}
        {/* ========================================================= */}
        {activeTab === 'weekly' && (
          <>
            <div className="grid lg:grid-cols-[1fr_2fr] gap-6 items-stretch">
              <Card title="Supervisor Details">
                <div className="flex flex-col gap-4">
                  <label className="flex flex-col gap-1 text-[11px] font-bold text-slate-500 uppercase">
                    Author Title
                    <input type="text" value={authorTitle} onChange={e => setAuthorTitle(e.target.value)} className="rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 outline-none font-bold text-slate-900 text-xs" />
                  </label>
                  <label className="flex flex-col gap-1 text-[11px] font-bold text-slate-500 uppercase">
                    Department
                    <input type="text" value={department} onChange={e => setDepartment(e.target.value)} className="rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 outline-none font-medium text-xs text-slate-900" />
                  </label>
                </div>
              </Card>
              <Card title="Executive Summary & Weekly KPIs (Summary of All)">
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 h-full">
                  <div className="bg-slate-50 p-3 rounded-xl border border-slate-100 shadow-inner flex flex-col justify-center"><NumberInput label="Arrived" value={containersArrived} onChange={setContainersArrived} /></div>
                  <div className="bg-slate-50 p-3 rounded-xl border border-slate-100 shadow-inner flex flex-col justify-center"><NumberInput label="Lead Time" value={avgLeadTime} onChange={setAvgLeadTime} /></div>
                  <div className="bg-slate-50 p-3 rounded-xl border border-slate-100 shadow-inner flex flex-col justify-center"><NumberInput label="Trucks %" value={truckAvailability} onChange={setTruckAvailability} suffix="%" /></div>
                  <div className="bg-slate-50 p-3 rounded-xl border border-slate-100 shadow-inner flex flex-col justify-center"><NumberInput label="Buffer" value={buffer} onChange={setBuffer} suffix="cntrs" /></div>
                </div>
              </Card>
            </div>

            <Card title="Next Vessels Monitoring" subtitle="Arrival Pipeline - Aggregated" action={
              <div className="flex items-center gap-2 no-print">
                <label className="flex items-center gap-2 px-3 py-1.5 bg-slate-100 text-slate-700 text-[11px] rounded-lg border border-slate-200 cursor-pointer hover:bg-slate-200 font-bold uppercase transition-all shadow-sm">
                  <span className="material-icons text-sm">upload</span>Upload<input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleVesselFileUpload} />
                </label>
                <button onClick={handleAddVessel} className="px-3 py-1.5 bg-slate-900 text-white text-[11px] rounded-lg shadow hover:bg-slate-800 transition-colors font-bold uppercase">Add Entry</button>
              </div>
            }>
               <div className="overflow-x-auto border border-slate-200 rounded-xl">
                 <table className="w-full text-[11px] min-w-[1100px] border-collapse">
                   <thead className="bg-[#f8fafc]">
                     <tr className="text-slate-500 border-b border-slate-200 uppercase text-[9px] font-black tracking-widest">
                       <th className="text-left py-3 px-4 w-[25%]">Vessel Name</th>
                       <th className="text-right py-3 pr-4 w-[10%]">Total Cntrs</th>
                       <th className="text-center py-3 w-[12%]">ETA</th>
                       <th className="text-center py-3 w-[12%] text-red-600">Deadline</th>
                       <th className="text-left py-3 px-2 w-[15%]">Warehouse</th>
                       <th className="text-right py-3 pr-4 w-[12%]">Delivery</th>
                       <th className="text-center py-3 pl-4 w-[12%]">Status</th>
                       <th className="w-8 no-print"></th>
                     </tr>
                   </thead>
                   <tbody className="bg-white">
                     {sortedVessels.map(v => (
                       <React.Fragment key={v.id}>
                         <tr className={`group border-b border-slate-100 cursor-pointer transition-all hover:bg-slate-50 ${expandedVesselId === v.id ? 'bg-slate-50 ring-1 ring-inset ring-slate-200' : ''}`} onClick={() => toggleVesselExpand(v.id)}>
                           <td className="py-2.5 px-4 flex items-center gap-2">
                             <span className={`material-icons text-[18px] text-slate-400 transition-transform ${expandedVesselId === v.id ? 'rotate-90 text-red-600' : ''}`}>chevron_right</span>
                             <div className="flex flex-col w-full">
                               <input className="w-full outline-none bg-transparent font-black text-slate-900 text-[13px] uppercase tracking-tight" value={v.vesselName} onClick={(e) => e.stopPropagation()} onChange={e => handleVesselChange(v.id, 'vesselName', e.target.value)} />
                               <input className="w-full outline-none bg-transparent text-[10px] text-slate-400 italic mt-0.5" placeholder="Add status details..." value={v.update} onClick={(e) => e.stopPropagation()} onChange={e => handleVesselChange(v.id, 'update', e.target.value)} />
                             </div>
                           </td>
                           <td className="text-right pr-4 font-black text-[15px] text-slate-900">{v.quantity}</td>
                           <td className="text-center"><input className="outline-none bg-transparent font-bold text-slate-600 text-xs" type="date" onClick={(e) => e.stopPropagation()} value={v.eta} onChange={e => handleVesselChange(v.id, 'eta', e.target.value)} /></td>
                           <td className="text-center"><input className="outline-none bg-transparent font-black text-red-600 text-xs" type="date" onClick={(e) => e.stopPropagation()} value={v.deadlineDate} onChange={e => handleVesselChange(v.id, 'deadlineDate', e.target.value)} /></td>
                           <td className="px-2"><input className="w-full text-[10px] outline-none bg-transparent text-slate-600 font-black uppercase" placeholder="WAREHOUSE" onClick={(e) => e.stopPropagation()} value={v.warehouse} onChange={e => handleVesselChange(v.id, 'warehouse', e.target.value)} /></td>
                           <td className="text-right pr-4"><input className="w-full text-right text-[10px] outline-none bg-transparent font-bold bg-amber-50 rounded px-1 border border-amber-100" type="date" onClick={(e) => e.stopPropagation()} value={v.deliveryScheduleDate} onChange={e => handleVesselChange(v.id, 'deliveryScheduleDate', e.target.value)} /></td>
                           <td className="text-center"><input className="w-full text-center outline-none bg-transparent font-black italic text-slate-900 uppercase" value={v.status} onClick={(e) => e.stopPropagation()} onChange={e => handleVesselChange(v.id, 'status', e.target.value)} /></td>
                           <td className="text-center no-print px-2"><button onClick={(e) => { e.stopPropagation(); handleRemoveVessel(v.id); }} className="text-slate-300 hover:text-red-500 opacity-0 group-hover:opacity-100"><span className="material-icons text-sm">delete</span></button></td>
                         </tr>
                         {expandedVesselId === v.id && (
                           <tr className="bg-slate-50/40">
                             <td colSpan={8} className="p-0 border-b border-slate-200">
                               <div className="mx-8 my-4 border border-slate-200 rounded-xl bg-white overflow-hidden shadow-md">
                                 <div className="bg-slate-50 px-4 py-2 border-b border-slate-200 flex justify-between items-center">
                                   <span className="text-[10px] font-black uppercase text-slate-500 tracking-wider">Individual BL Breakdown</span>
                                   <button onClick={() => handleAddVesselDetail(v.id)} className="px-2 py-1 bg-white border border-slate-200 text-[9px] font-black rounded hover:bg-slate-50 transition-colors uppercase flex items-center gap-1 shadow-sm"><span className="material-icons text-[10px]">add</span> Add BL</button>
                                 </div>
                                 <div className="overflow-x-auto">
                                   <table className="w-full text-[10px] border-collapse min-w-[900px]">
                                     <thead className="bg-white border-b border-slate-100">
                                       <tr className="text-slate-400 font-black uppercase text-[8px] tracking-widest">
                                         <th className="py-2.5 px-4 text-left w-[18%]">BL NUMBER</th>
                                         <th className="py-2.5 px-4 text-left w-[15%]">PO NUMBER</th>
                                         <th className="py-2.5 px-4 text-left w-[12%]">PO TOTAL BATCH</th>
                                         <th className="py-2.5 px-4 text-left w-[15%]">CARGO TYPE</th>
                                         <th className="py-2.5 px-4 text-right w-[8%]">QTY</th>
                                         <th className="py-2.5 px-4 text-center w-[15%]">STATUS</th>
                                         <th className="py-2.5 px-4 text-left w-[12%]">BROKER</th>
                                         <th className="w-8"></th>
                                       </tr>
                                     </thead>
                                     <tbody className="divide-y divide-slate-50">
                                       {v.details?.map((detail) => (
                                         <tr key={detail.id} className="hover:bg-blue-50/30 transition-colors group/detail">
                                           <td className="py-2 px-4"><input className="w-full bg-transparent outline-none font-black text-slate-800" value={detail.blNumber} onChange={e => handleVesselDetailChange(v.id, detail.id, 'blNumber', e.target.value)} /></td>
                                           <td className="py-2 px-4"><input className="w-full bg-transparent outline-none font-semibold text-slate-600" value={detail.po} onChange={e => handleVesselDetailChange(v.id, detail.id, 'po', e.target.value)} /></td>
                                           <td className="py-2 px-4"><input className="w-full bg-transparent outline-none font-medium text-slate-500" value={detail.poTotalBatch} onChange={e => handleVesselDetailChange(v.id, detail.id, 'poTotalBatch', e.target.value)} /></td>
                                           <td className="py-2 px-4"><input className="w-full bg-transparent outline-none font-medium text-slate-600 italic" value={detail.cargoType} onChange={e => handleVesselDetailChange(v.id, detail.id, 'cargoType', e.target.value)} /></td>
                                           <td className="py-2 px-4 text-right bg-slate-50/50"><input className="w-full text-right bg-transparent outline-none font-black text-slate-900" type="number" value={detail.qty} onChange={e => handleVesselDetailChange(v.id, detail.id, 'qty', Number(e.target.value))} /></td>
                                           <td className="py-2 px-4 text-center"><input className="w-full text-center bg-transparent outline-none font-bold text-slate-500 italic uppercase" value={detail.status} onChange={e => handleVesselDetailChange(v.id, detail.id, 'status', e.target.value)} /></td>
                                           <td className="py-2 px-4"><input className="w-full bg-transparent outline-none font-bold text-slate-700 uppercase text-[9px]" value={detail.broker} onChange={e => handleVesselDetailChange(v.id, detail.id, 'broker', e.target.value)} /></td>
                                           <td className="text-center px-2 no-print"><button onClick={() => handleRemoveVesselDetail(v.id, detail.id)} className="text-slate-200 hover:text-red-500 opacity-0 group-hover/detail:opacity-100"><span className="material-icons text-[14px]">remove_circle_outline</span></button></td>
                                         </tr>
                                       ))}
                                     </tbody>
                                   </table>
                                 </div>
                               </div>
                             </td>
                           </tr>
                         )}
                       </React.Fragment>
                     ))}
                   </tbody>
                 </table>
               </div>
            </Card>

            <Card title="Quota Situation Monitoring" subtitle="Round-based Import License Tracker" action={
              <div className="flex items-center gap-3 no-print">
                <label className="flex items-center gap-2 px-3 py-1.5 bg-slate-100 text-slate-700 text-[11px] rounded-lg border border-slate-200 cursor-pointer hover:bg-slate-200 font-bold uppercase shadow-sm">
                  <span className="material-icons text-sm">upload_file</span>Upload<input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleQuotaFileUpload} />
                </label>
                <button onClick={handleAddQuota} className="px-3 py-1.5 bg-red-600 text-white text-[11px] rounded-lg shadow hover:bg-red-700 font-bold uppercase">Add New BL</button>
              </div>
            }>
               <div className="grid grid-cols-2 gap-8 mb-6 mt-2">
                  <div className="flex flex-col gap-2">
                     <div className="bg-yellow-400 text-center font-black text-xs py-1.5 border border-slate-300 rounded-t-lg">EV QUOTA BALANCE</div>
                     <div className="flex flex-col gap-1.5 bg-slate-50 border border-slate-200 rounded-b-lg p-3 text-[11px] font-black uppercase">
                        <div className="flex justify-between items-center text-slate-500"><span>Initial Balance</span><div className="flex items-center gap-1"><span className="text-slate-900 font-mono">$</span><input type="number" className="bg-white border px-2 py-1 outline-none text-right w-40 font-mono text-xs rounded shadow-sm" value={evBalance} onChange={e => setEvBalance(Number(e.target.value))} /></div></div>
                        <div className="flex justify-between items-center text-red-600 border-t border-slate-200 pt-1.5 mt-1"><span>Total Quota Used (FOB)</span><span className="font-mono text-xs">-${evTotalSpent.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span></div>
                        <div className="flex justify-between items-center text-emerald-700 border-t-2 border-slate-300 pt-1.5 mt-1 bg-emerald-50 px-2 py-1 -mx-3 rounded-b-lg"><span>Final Remaining Balance</span><span className="font-mono text-xs">${(evBalance - evTotalSpent).toLocaleString('en-US', { minimumFractionDigits: 2 })}</span></div>
                     </div>
                  </div>
                  <div className="flex flex-col gap-2">
                     <div className="bg-yellow-400 text-center font-black text-xs py-1.5 border border-slate-300 rounded-t-lg">PHEV QUOTA BALANCE</div>
                     <div className="flex flex-col gap-1.5 bg-slate-50 border border-slate-200 rounded-b-lg p-3 text-[11px] font-black uppercase">
                        <div className="flex justify-between items-center text-slate-500"><span>Initial Balance</span><div className="flex items-center gap-1"><span className="text-slate-900 font-mono">$</span><input type="number" className="bg-white border px-2 py-1 outline-none text-right w-40 font-mono text-xs rounded shadow-sm" value={phevBalance} onChange={e => setPhevBalance(Number(e.target.value))} /></div></div>
                        <div className="flex justify-between items-center text-red-600 border-t border-slate-200 pt-1.5 mt-1"><span>Total Quota Used (FOB)</span><span className="font-mono text-xs">-${phevTotalSpent.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span></div>
                        <div className="flex justify-between items-center text-emerald-700 border-t-2 border-slate-300 pt-1.5 mt-1 bg-emerald-50 px-2 py-1 -mx-3 rounded-b-lg"><span>Final Remaining Balance</span><span className="font-mono text-xs">${(phevBalance - phevTotalSpent).toLocaleString('en-US', { minimumFractionDigits: 2 })}</span></div>
                     </div>
                  </div>
               </div>
               <div className="grid lg:grid-cols-2 gap-10">
                  <div><RoundTable type="EV" round={1} /><RoundTable type="EV" round={2} /><RoundTable type="EV" round={3} /><RoundTable type="EV" round={4} /><div className="mt-4 p-3 bg-slate-900 text-white rounded-xl shadow flex justify-between items-center"><span className="font-black uppercase text-[10px] tracking-widest">Total EV Pipeline</span><div className="flex gap-6 font-mono text-xs"><span className="text-yellow-400 font-black">{evTotalQty} Units</span><span className="font-black text-sm">${evTotalSpent.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span></div></div></div>
                  <div><RoundTable type="PHEV" round={1} /><RoundTable type="PHEV" round={2} /><RoundTable type="PHEV" round={3} /><RoundTable type="PHEV" round={4} /><div className="mt-4 p-3 bg-slate-900 text-white rounded-xl shadow flex justify-between items-center"><span className="font-black uppercase text-[10px] tracking-widest">Total PHEV Pipeline</span><div className="flex gap-6 font-mono text-xs"><span className="text-yellow-400 font-black">{phevTotalQty} Units</span><span className="font-black text-sm">${phevTotalSpent.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span></div></div></div>
               </div>
            </Card>

            <div className="grid lg:grid-cols-3 gap-6">
               <Card title="1. Executive Summary"><TextAreaField label="" value={executiveSummary} onChange={setExecutiveSummary} minRows={6}/></Card>
               <Card title="2. Highlights"><TextAreaField label="" value={highlights} onChange={setHighlights} minRows={6}/></Card>
               <Card title="3. Issues & Next Steps"><TextAreaField label="Current Issues" value={issues} onChange={setIssues} minRows={2}/><TextAreaField label="Action Plan" value={nextActions} onChange={setNextActions} minRows={2}/></Card>
            </div>
          </>
        )}

        {/* ========================================================= */}
        {/* TAB 2: LOGISTICS (Write entire situation for director)    */}
        {/* ========================================================= */}
        {activeTab === 'logistics' && (
          <div className="flex flex-col gap-6">
            <Card title="Logistics Department Situation Report" subtitle={`Prepared by ${author} for ${directorName}`}>
              <div className="flex flex-col gap-4">
                <p className="text-xs text-slate-500">
                  Write or edit the complete logistics situation report below. This formatted briefing is structured specifically to send directly to the Director.
                </p>
                <TextAreaField 
                  label="Full Logistics Situation Report" 
                  value={logisticsDirectorReport} 
                  onChange={setLogisticsDirectorReport} 
                  minRows={18} 
                />
              </div>
            </Card>

            <div className="grid md:grid-cols-3 gap-4">
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase">Containers Arrived</span>
                <span className="text-xl font-black text-slate-900 font-mono">{containersArrived}</span>
              </div>
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase">Average Lead Time</span>
                <span className="text-xl font-black text-slate-900 font-mono">{avgLeadTime} days</span>
              </div>
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col gap-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase">Truck Fleet Availability</span>
                <span className="text-xl font-black text-slate-900 font-mono">{truckAvailability}%</span>
              </div>
            </div>
          </div>
        )}

        {/* ========================================================= */}
        {/* TAB 3: BROKER (Write entire situation for director)       */}
        {/* ========================================================= */}
        {activeTab === 'broker' && (
          <div className="flex flex-col gap-6">
            <Card title="Customs Broker Department Situation Report" subtitle={`Prepared by ${author} for ${directorName}`}>
              <div className="flex flex-col gap-4">
                <p className="text-xs text-slate-500">
                  Write or edit the complete Customs Broker situation report below to send to the Director.
                </p>
                <TextAreaField 
                  label="Full Customs Broker Situation Report" 
                  value={brokerDirectorReport} 
                  onChange={setBrokerDirectorReport} 
                  minRows={18} 
                />
              </div>
            </Card>

            <Card title="Broker Performance Quick Reference">
              <div className="overflow-x-auto border border-slate-200 rounded-xl">
                <table className="w-full text-xs">
                  <thead className="bg-slate-50 border-b border-slate-200">
                    <tr className="text-[10px] font-bold uppercase text-slate-500">
                      <th className="text-left p-3">Broker Entity</th>
                      <th className="text-right p-3">Volume</th>
                      <th className="text-right p-3">Avg Days</th>
                      <th className="text-center p-3">Green %</th>
                      <th className="text-center p-3">Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {brokers.map(b => (
                      <tr key={b.id} className="hover:bg-slate-50">
                        <td className="p-3 font-bold text-slate-900">{b.name}</td>
                        <td className="text-right p-3 font-mono">{b.processes}</td>
                        <td className="text-right p-3 font-mono">{b.avgClearanceDays}</td>
                        <td className="text-center p-3 font-mono">{b.greenChannelRate}%</td>
                        <td className="text-center p-3"><span className={`text-[10px] font-bold px-2 py-0.5 rounded-md border ${statusPillClasses[b.status]}`}>{b.status}</span></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Card>
          </div>
        )}

        {/* ========================================================= */}
        {/* TAB 4: FINANCIAL (Write entire situation for director)    */}
        {/* ========================================================= */}
        {activeTab === 'financial' && (
          <div className="flex flex-col gap-6">
            <Card title="Financial Department Situation Report" subtitle={`Prepared by ${author} for ${directorName}`}>
              <div className="flex flex-col gap-4">
                <p className="text-xs text-slate-500">
                  Write or edit the complete Financial situation report below to send to the Director.
                </p>
                <TextAreaField 
                  label="Full Financial Situation Report" 
                  value={financialDirectorReport} 
                  onChange={setFinancialDirectorReport} 
                  minRows={18} 
                />
              </div>
            </Card>

            <div className="grid md:grid-cols-2 gap-6">
              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col gap-2">
                <span className="text-xs font-black text-slate-900 uppercase">EV Quota Budget Status</span>
                <div className="flex justify-between text-xs text-slate-500"><span>Initial:</span><span className="font-mono">${evBalance.toLocaleString()}</span></div>
                <div className="flex justify-between text-xs text-red-600 font-semibold"><span>Spent:</span><span className="font-mono">-${evTotalSpent.toLocaleString()}</span></div>
                <div className="flex justify-between text-xs text-emerald-700 font-black border-t pt-2"><span>Remaining Balance:</span><span className="font-mono">${(evBalance - evTotalSpent).toLocaleString()}</span></div>
              </div>
              <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col gap-2">
                <span className="text-xs font-black text-slate-900 uppercase">PHEV Quota Budget Status</span>
                <div className="flex justify-between text-xs text-slate-500"><span>Initial:</span><span className="font-mono">${phevBalance.toLocaleString()}</span></div>
                <div className="flex justify-between text-xs text-red-600 font-semibold"><span>Spent:</span><span className="font-mono">-${phevTotalSpent.toLocaleString()}</span></div>
                <div className="flex justify-between text-xs text-emerald-700 font-black border-t pt-2"><span>Remaining Balance:</span><span className="font-mono">${(phevBalance - phevTotalSpent).toLocaleString()}</span></div>
              </div>
            </div>
          </div>
        )}

      </main>
    </div>
  );
};

export default App;
