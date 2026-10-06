
export interface BrokerRow {
  id: string;
  name: string;
  processes: number;
  avgClearanceDays: number;
  greenChannelRate: number;
  issues: string;
  status: "OK" | "Attention" | "Critical";
}

export interface TruckSupplierRow {
  id: string;
  supplierName: string;
  allocationRate: number;
  onTimeRate: number;
  comments: string;
}

export interface VesselDetail {
  id: string;
  blNumber: string;
  po: string;
  poTotalBatch?: string;
  voyage: string;
  qty: number;
  cargoType: string;
  status: string;
  broker?: string;
}

export interface VesselRow {
  id: string;
  vesselName: string;
  quantity: number;
  eta: string;
  status: string;
  update: string;
  warehouse: string;
  deliveryScheduleDate: string;
  deadlineDate: string;
  details?: VesselDetail[];
}

export interface LIRow {
  id: string;
  liNumber: string;
  poNumber: string;
  model: string;
  status: string;
}

export interface QuotaEntry {
  id: string;
  vessel: string;
  po: string;
  ncm: string;
  type: string; // EV, PHEV
  round: number; // 1, 2, 3, 4
  li: string;
  registerDate: string;
  fobValue: number;
  qty: number;
  exigenciaDate: string;
  docsDate: string;
  status: string;
  di: string;
  channel: string;
}

export type PaymentCategory = "Warehouse" | "Truck Carrier" | "Broker" | "Other Invoices";

export interface PaymentRow {
  id: string;
  category: PaymentCategory;
  entity: string;
  description: string;
  plannedAmount: number;
  pendingAmount: number;
  currency: "BRL" | "USD" | "CNY";
  status: "Paid" | "Scheduled" | "Hold";
}
