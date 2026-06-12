/**
 * Benchmark dataset — one deterministic mock company, two business views.
 *
 * Everything downstream (billing workbook, cost-center workbook, gold
 * liquidity report) derives from this module with a fixed seed, so the
 * gold answers never need to see the generated workbooks and the blind
 * agent can be tested against any seed.
 */

export type Month = string; // '2026-01' … ISO year-month

export const MONTHS: Month[] = ['2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06'];
export const FORECAST_MONTHS: Month[] = ['2026-07', '2026-08', '2026-09'];

export interface Customer {
  id: string;
  name: string;
  tier: 'Enterprise' | 'Mid-Market' | 'SMB';
  contractType: 'Subscription' | 'Project' | 'Hybrid';
}

export interface Invoice {
  id: string;
  customerId: string;
  issueMonth: Month;
  dueMonth: Month;
  category: string;
  amount: number;
  status: 'Paid' | 'Partial' | 'Open' | 'Overdue';
  memo?: string;
}

export interface Payment {
  id: string;
  invoiceId: string;
  customerId: string;
  month: Month;
  amount: number;
}

export interface Credit {
  id: string;
  customerId: string;
  invoiceId: string;
  month: Month;
  amount: number; // positive number; reduces cash collected in `month`
  reason: string;
}

export interface DeferredContract {
  invoiceId: string;
  customerId: string;
  total: number;
  startMonth: Month;
  monthlyRecognition: number;
}

export interface Department {
  code: string;
  name: string;
}

export interface CostCenter {
  code: string;
  name: string;
  dept: string;
  payrollShare: number; // shares per dept sum to 1
}

export interface Vendor {
  id: string;
  name: string;
  dept: string | 'SHARED';
  costCenter?: string;
  category: string;
}

export interface VendorSpendRow {
  vendorId: string;
  month: Month;
  amount: number;
}

export interface PayrollRow {
  dept: string;
  month: Month;
  amount: number;
  headcount: number;
}

export interface CapexItem {
  id: string;
  item: string;
  vendorId: string;
  dept: string;
  costCenter: string;
  month: Month;
  amount: number;
}

export interface Dataset {
  seed: number;
  months: Month[];
  forecastMonths: Month[];
  customers: Customer[];
  billingCategories: { name: string; revenueType: 'Recurring' | 'One-time'; code: string }[];
  invoices: Invoice[];
  payments: Payment[];
  credits: Credit[];
  deferred: DeferredContract[];
  departments: Department[];
  costCenters: CostCenter[];
  vendors: Vendor[];
  vendorSpend: VendorSpendRow[];
  payroll: PayrollRow[];
  capex: CapexItem[];
  /** Hidden mapping: department → share of SHARED vendor spend (sums to 1). */
  allocShares: Record<string, number>;
  /** Forecast (Jul–Sep) dept totals — context noise, excluded from actuals. */
  forecast: { dept: string; month: Month; amount: number }[];
}

/** mulberry32 — small deterministic PRNG. */
export function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const r2 = (n: number): number => Math.round(n * 100) / 100;
const pick = <T>(r: () => number, arr: T[]): T => arr[Math.floor(r() * arr.length)];

function addMonths(m: Month, k: number): Month {
  const [y, mo] = m.split('-').map(Number);
  const idx = y * 12 + (mo - 1) + k;
  const ny = Math.floor(idx / 12);
  const nm = (idx % 12) + 1;
  return `${ny}-${String(nm).padStart(2, '0')}`;
}

export function generateDataset(seed: number): Dataset {
  const r = rng(seed);

  const customerNames = [
    'Acme Industrial',
    'Borealis Health',
    'Cobalt Retail',
    'Dyno Logistics',
    'Everfield Energy',
    'Fjord Analytics',
    'Granite Foods',
    'Helio Media'
  ];
  const tiers: Customer['tier'][] = ['Enterprise', 'Enterprise', 'Mid-Market', 'Mid-Market', 'Mid-Market', 'SMB', 'SMB', 'SMB'];
  const contractTypes: Customer['contractType'][] = ['Subscription', 'Hybrid', 'Subscription', 'Project', 'Hybrid', 'Subscription', 'Project', 'Subscription'];
  const customers: Customer[] = customerNames.map((name, i) => ({
    id: `C${String(i + 1).padStart(2, '0')}`,
    name,
    tier: tiers[i],
    contractType: contractTypes[i]
  }));

  const billingCategories: Dataset['billingCategories'] = [
    { name: 'SaaS Subscription', revenueType: 'Recurring', code: 'REV-SAAS' },
    { name: 'Professional Services', revenueType: 'One-time', code: 'REV-PS' },
    { name: 'Support Plan', revenueType: 'Recurring', code: 'REV-SUP' },
    { name: 'Hardware', revenueType: 'One-time', code: 'REV-HW' },
    { name: 'Training', revenueType: 'One-time', code: 'REV-TRN' }
  ];
  const catMult: Record<string, number> = {
    'SaaS Subscription': 1.0,
    'Professional Services': 1.4,
    'Support Plan': 0.35,
    Hardware: 0.9,
    Training: 0.25
  };
  const tierBase: Record<Customer['tier'], number> = { Enterprise: 12000, 'Mid-Market': 6000, SMB: 2500 };

  const invoices: Invoice[] = [];
  const payments: Payment[] = [];
  const deferred: DeferredContract[] = [];
  let invSeq = 1000;
  let pmtSeq = 2000;

  for (const customer of customers) {
    for (const month of MONTHS) {
      const count = 1 + (r() < 0.35 ? 1 : 0);
      for (let k = 0; k < count; k++) {
        const category =
          customer.contractType === 'Subscription' && k === 0
            ? 'SaaS Subscription'
            : pick(r, billingCategories).name;
        const amount = r2(tierBase[customer.tier] * catMult[category] * (0.8 + 0.4 * r()));
        const roll = r();
        const status: Invoice['status'] = roll < 0.62 ? 'Paid' : roll < 0.77 ? 'Partial' : roll < 0.89 ? 'Open' : 'Overdue';
        const inv: Invoice = {
          id: `INV-${invSeq++}`,
          customerId: customer.id,
          issueMonth: month,
          dueMonth: addMonths(month, 1),
          category,
          amount,
          status
        };
        invoices.push(inv);
        const clampMonth = (m: Month): Month => (MONTHS.includes(m) ? m : MONTHS[MONTHS.length - 1]);
        if (status === 'Paid') {
          const payMonth = clampMonth(addMonths(month, r() < 0.25 ? 1 : 0));
          payments.push({ id: `PMT-${pmtSeq++}`, invoiceId: inv.id, customerId: customer.id, month: payMonth, amount });
        } else if (status === 'Partial') {
          const first = r2(amount * 0.55);
          const second = r2(amount * 0.3);
          payments.push({ id: `PMT-${pmtSeq++}`, invoiceId: inv.id, customerId: customer.id, month, amount: first });
          payments.push({ id: `PMT-${pmtSeq++}`, invoiceId: inv.id, customerId: customer.id, month: clampMonth(addMonths(month, 1)), amount: second });
        }
      }
    }
  }

  // Annual prepaid contracts (deferred revenue): cash lands in January.
  for (const ci of [0, 3]) {
    const customer = customers[ci];
    const total = ci === 0 ? 60000 : 48000;
    const inv: Invoice = {
      id: `INV-${invSeq++}`,
      customerId: customer.id,
      issueMonth: MONTHS[0],
      dueMonth: MONTHS[0],
      category: 'SaaS Subscription',
      amount: total,
      status: 'Paid',
      memo: 'Annual prepaid contract'
    };
    invoices.push(inv);
    payments.push({ id: `PMT-${pmtSeq++}`, invoiceId: inv.id, customerId: customer.id, month: MONTHS[0], amount: total });
    deferred.push({
      invoiceId: inv.id,
      customerId: customer.id,
      total,
      startMonth: MONTHS[0],
      monthlyRecognition: r2(total / 6)
    });
  }

  // Credits applied against paid invoices.
  const credits: Credit[] = [];
  const reasons = ['Service credit', 'Billing error', 'Goodwill', 'SLA breach'];
  const paidInvoices = invoices.filter((i) => i.status === 'Paid' && !i.memo);
  for (let k = 0; k < 6; k++) {
    const inv = paidInvoices[Math.floor(r() * paidInvoices.length)];
    const paymentMonth = payments.find((p) => p.invoiceId === inv.id)!.month;
    let month = addMonths(paymentMonth, r() < 0.4 ? 1 : 0);
    if (!MONTHS.includes(month)) month = MONTHS[MONTHS.length - 1];
    credits.push({
      id: `CR-${String(k + 1).padStart(2, '0')}`,
      customerId: inv.customerId,
      invoiceId: inv.id,
      month,
      amount: r2(inv.amount * (0.05 + 0.07 * r())),
      reason: reasons[k % reasons.length]
    });
  }

  const departments: Department[] = [
    { code: 'ENG', name: 'Engineering' },
    { code: 'SAL', name: 'Sales' },
    { code: 'MKT', name: 'Marketing' },
    { code: 'OPS', name: 'Operations' },
    { code: 'GA', name: 'General & Admin' }
  ];
  const costCenters: CostCenter[] = [
    { code: 'ENG-100', name: 'Platform', dept: 'ENG', payrollShare: 0.65 },
    { code: 'ENG-110', name: 'Quality', dept: 'ENG', payrollShare: 0.35 },
    { code: 'SAL-200', name: 'Field Sales', dept: 'SAL', payrollShare: 0.7 },
    { code: 'SAL-210', name: 'Sales Ops', dept: 'SAL', payrollShare: 0.3 },
    { code: 'MKT-300', name: 'Demand Gen', dept: 'MKT', payrollShare: 0.6 },
    { code: 'MKT-310', name: 'Brand', dept: 'MKT', payrollShare: 0.4 },
    { code: 'OPS-400', name: 'Fulfilment', dept: 'OPS', payrollShare: 0.55 },
    { code: 'OPS-410', name: 'Logistics', dept: 'OPS', payrollShare: 0.45 },
    { code: 'GA-500', name: 'Finance', dept: 'GA', payrollShare: 0.6 },
    { code: 'GA-520', name: 'People & Places', dept: 'GA', payrollShare: 0.4 }
  ];

  const vendors: Vendor[] = [
    { id: 'V01', name: 'CloudNine Hosting', dept: 'SHARED', category: 'Cloud Infrastructure' },
    { id: 'V02', name: 'OfficeWorks', dept: 'SHARED', category: 'Facilities' },
    { id: 'V03', name: 'TalentBridge Recruiting', dept: 'GA', costCenter: 'GA-520', category: 'Recruiting' },
    { id: 'V04', name: 'AdSpark', dept: 'MKT', costCenter: 'MKT-300', category: 'Advertising' },
    { id: 'V05', name: 'FreightCo', dept: 'OPS', costCenter: 'OPS-410', category: 'Freight' },
    { id: 'V06', name: 'LegalEase', dept: 'GA', costCenter: 'GA-500', category: 'Legal' },
    { id: 'V07', name: 'DataDome Analytics', dept: 'ENG', costCenter: 'ENG-100', category: 'Analytics Tools' },
    { id: 'V08', name: 'SecureNet', dept: 'ENG', costCenter: 'ENG-110', category: 'Security' },
    { id: 'V09', name: 'CafePlus Catering', dept: 'OPS', costCenter: 'OPS-400', category: 'Food Services' },
    { id: 'V10', name: 'ToolForge Licenses', dept: 'SAL', costCenter: 'SAL-210', category: 'Software Licenses' }
  ];
  const vendorBase: Record<string, number> = {
    V01: 9000, V02: 6500, V03: 4200, V04: 5200, V05: 3800,
    V06: 2400, V07: 3100, V08: 2800, V09: 1500, V10: 2100
  };

  const vendorSpend: VendorSpendRow[] = [];
  for (const vendor of vendors) {
    for (const month of MONTHS) {
      let amount = vendorBase[vendor.id] * (0.85 + 0.3 * r());
      if (vendor.id === 'V04' && month === '2026-04') amount *= 2.6; // campaign spike anomaly
      vendorSpend.push({ vendorId: vendor.id, month, amount: r2(amount) });
    }
  }

  const payrollBase: Record<string, number> = { ENG: 95000, SAL: 60000, MKT: 38000, OPS: 45000, GA: 30000 };
  const payroll: PayrollRow[] = [];
  departments.forEach((dept) => {
    MONTHS.forEach((month, mi) => {
      const ramp = dept.code === 'ENG' ? 1500 * mi : 0;
      const amount = r2((payrollBase[dept.code] + ramp) * (0.97 + 0.06 * r()));
      payroll.push({ dept: dept.code, month, amount, headcount: Math.round(amount / 9500) });
    });
  });

  const capex: CapexItem[] = [
    { id: 'CAP-01', item: 'Server cluster expansion', vendorId: 'V08', dept: 'ENG', costCenter: 'ENG-100', month: '2026-02', amount: 24000 },
    { id: 'CAP-02', item: 'Warehouse forklift', vendorId: 'V05', dept: 'OPS', costCenter: 'OPS-410', month: '2026-04', amount: 18000 },
    { id: 'CAP-03', item: 'Office buildout — east wing', vendorId: 'V02', dept: 'GA', costCenter: 'GA-520', month: '2026-05', amount: 15000 }
  ];

  const allocShares: Record<string, number> = { ENG: 0.4, SAL: 0.15, MKT: 0.15, OPS: 0.2, GA: 0.1 };

  const forecast: Dataset['forecast'] = [];
  for (const dept of departments) {
    const lastPayroll = payroll.filter((p) => p.dept === dept.code).slice(-1)[0].amount;
    FORECAST_MONTHS.forEach((month, k) => {
      forecast.push({ dept: dept.code, month, amount: r2(lastPayroll * Math.pow(1.03, k + 1) * 1.45) });
    });
  }

  return {
    seed,
    months: MONTHS,
    forecastMonths: FORECAST_MONTHS,
    customers,
    billingCategories,
    invoices,
    payments,
    credits,
    deferred,
    departments,
    costCenters,
    vendors,
    vendorSpend,
    payroll,
    capex,
    allocShares,
    forecast
  };
}

// pharos:eof
