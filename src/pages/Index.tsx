import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  Activity,
  BadgeCheck,
  BriefcaseBusiness,
  CalendarDays,
  CheckCircle2,
  ClipboardList,
  Copy,
  Download,
  FileCheck2,
  FileText,
  Hotel,
  LayoutDashboard,
  Lock,
  LogIn,
  LogOut,
  Mail,
  MapPin,
  Menu,
  Plane,
  Plus,
  Printer,
  RefreshCcw,
  Save,
  Search,
  Send,
  ShieldCheck,
  Trash2,
  Upload,
  UserRound,
  WalletCards,
} from "lucide-react";
import { toast } from "sonner";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Textarea } from "@/components/ui/textarea";
import { cn } from "@/lib/utils";

type Role = "Admin" | "Documentation Officer" | "Reviewer";
type Status = "Draft" | "Review" | "Verified" | "Ready" | "Submitted";
type TemplateKey = "financial" | "employment" | "salary" | "itinerary" | "accommodation" | "cover";

const ROLE_PERMS: Record<Role, { verify: boolean; submit: boolean; setStatus: Status[]; manageClients: boolean }> = {
  Admin: { verify: true, submit: true, setStatus: ["Draft", "Review", "Verified", "Ready", "Submitted"], manageClients: true },
  "Documentation Officer": { verify: false, submit: true, setStatus: ["Draft", "Review", "Ready", "Submitted"], manageClients: true },
  Reviewer: { verify: true, submit: false, setStatus: ["Review", "Verified", "Ready"], manageClients: false },
};
const can = (role: Role | undefined, action: "verify" | "submit" | "manageClients") => !!role && ROLE_PERMS[role][action];
const STATUS_FLOW: Status[] = ["Draft", "Review", "Verified", "Ready", "Submitted"];

type Client = {
  id: string;
  name: string;
  passport: string;
  nationality: string;
  dob: string;
  phone: string;
  email: string;
  address: string;
  visaCountry: string;
  visaType: string;
  notes: string;
};

type TransactionRow = { id: string; date: string; description: string; reference: string; debit: number; credit: number; balance: number };
type DayPlan = { id: string; date: string; city: string; plan: string };
type ReviewNote = { id: string; author: string; role: Role; text: string; at: string };
type DocumentDraft = {
  id: string;
  clientId: string;
  template: TemplateKey;
  title: string;
  status: Status;
  verified: boolean;
  watermark: boolean;
  updatedAt: string;
  fields: Record<string, string | number | boolean>;
  transactions?: TransactionRow[];
  dayPlans?: DayPlan[];
  reviewNotes?: ReviewNote[];
};
type ActivityLog = { id: string; text: string; at: string; icon: TemplateKey | "client" | "save" | "status" | "note" | "import" | "export"; actor?: string; role?: Role };
type Session = { name: string; role: Role } | null;

type Store = {
  clients: Client[];
  documents: DocumentDraft[];
  templates: { key: TemplateKey; name: string }[];
  transactions: TransactionRow[];
  activityLogs: ActivityLog[];
  session: Session;
};

const templateMeta: Record<TemplateKey, { name: string; short: string; icon: typeof FileText }> = {
  financial: { name: "Financial Evidence Summary", short: "Financial", icon: WalletCards },
  employment: { name: "Employment Certificate / Job Letter", short: "Employment", icon: BriefcaseBusiness },
  salary: { name: "Salary Certificate", short: "Salary", icon: BadgeCheck },
  itinerary: { name: "Travel Itinerary", short: "Itinerary", icon: Plane },
  accommodation: { name: "Accommodation / Hotel Booking Summary", short: "Accommodation", icon: Hotel },
  cover: { name: "Cover Letter / Visa Explanation Letter", short: "Cover Letter", icon: Mail },
};

const statusClass: Record<Status, string> = {
  Draft: "bg-muted text-muted-foreground border-border",
  Review: "bg-warning/15 text-foreground border-warning/30",
  Verified: "bg-success/15 text-success border-success/30",
  Ready: "bg-ready/15 text-ready border-ready/30",
  Submitted: "bg-submitted/15 text-submitted border-submitted/30",
};

const uid = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`;
const today = () => new Date().toISOString().slice(0, 10);
const fmtMoney = (value: number | string, currency = "SGD") =>
  `${currency} ${Number(value || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

const demoClient: Client = {
  id: "client-demo",
  name: "Ariana Tan",
  passport: "K1234567",
  nationality: "Singaporean",
  dob: "1991-08-16",
  phone: "+65 9123 4567",
  email: "ariana.tan@example.com",
  address: "18 Marina View, Singapore 018960",
  visaCountry: "Japan",
  visaType: "Tourist Visa",
  notes: "Family holiday application. Originals checked at intake appointment.",
};

const sampleTransactions: TransactionRow[] = [
  { id: uid(), date: "2026-01-02", description: "Opening balance", reference: "OB", debit: 0, credit: 0, balance: 18250 },
  { id: uid(), date: "2026-01-05", description: "Salary credit", reference: "PAY-0105", debit: 0, credit: 6200, balance: 24450 },
  { id: uid(), date: "2026-01-12", description: "Card payment", reference: "CARD-771", debit: 820, credit: 0, balance: 23630 },
  { id: uid(), date: "2026-01-18", description: "Travel savings transfer", reference: "TRF-220", debit: 0, credit: 1500, balance: 25130 },
  { id: uid(), date: "2026-01-29", description: "Utilities and rent", reference: "GIRO-448", debit: 3100, credit: 0, balance: 22030 },
];

const sampleDayPlans: DayPlan[] = [
  { id: uid(), date: "2026-04-03", city: "Tokyo", plan: "Arrival, hotel check-in, Shinjuku area visit" },
  { id: uid(), date: "2026-04-04", city: "Tokyo", plan: "Cultural sightseeing and client-funded leisure activities" },
  { id: uid(), date: "2026-04-05", city: "Kyoto", plan: "Train transfer, temple district visit" },
];

const defaultFields = (template: TemplateKey, client = demoClient): DocumentDraft["fields"] => {
  const common = { clientName: client.name, applicantName: client.name, issueDate: today() };
  const fields: Record<TemplateKey, DocumentDraft["fields"]> = {
    financial: {
      ...common,
      accountHolderName: client.name,
      accountNumber: "**** 7392",
      maskAccount: true,
      bankName: "Client Provided Bank",
      branch: "Main Branch",
      statementPeriod: "01 Jan 2026 - 31 Jan 2026",
      currency: "SGD",
      openingBalance: 18250,
      closingBalance: 22030,
      totalDebit: 3920,
      totalCredit: 7700,
      averageBalance: 22698,
    },
    employment: {
      ...common,
      companyName: "Northstar Analytics Pte. Ltd.",
      employeeName: client.name,
      designation: "Operations Manager",
      joiningDate: "2019-06-10",
      salary: "SGD 6,200 per month",
      hrPerson: "Maya Lim, HR Manager",
      companyAddress: "80 Robinson Road, Singapore 068898",
      bodyText:
        "This letter confirms that the above-named employee is currently employed with our company on a full-time basis. The employee has been granted leave for the proposed travel period and is expected to resume duties after the trip.",
    },
    salary: {
      ...common,
      employeeName: client.name,
      designation: "Operations Manager",
      monthlySalary: "6200",
      allowance: "450",
      netSalary: "6650",
      companyName: "Northstar Analytics Pte. Ltd.",
      authorizedPerson: "Maya Lim",
    },
    itinerary: {
      ...common,
      country: client.visaCountry,
      travelDates: "03 Apr 2026 - 12 Apr 2026",
      flightRoute: "Singapore → Tokyo → Kyoto → Singapore",
      hotelCity: "Tokyo / Kyoto",
      purpose: "Tourism and family leisure travel",
    },
    accommodation: {
      ...common,
      hotelName: "Sakura Central Hotel",
      address: "1-2-3 Shinjuku, Tokyo, Japan",
      checkIn: "2026-04-03",
      checkOut: "2026-04-08",
      bookingReference: "SAMPLE-HTL-90421",
      contact: "+81 3 0000 0000",
      notes: "Booking details are prepared for file review and must be verified against reservation confirmation.",
    },
    cover: {
      ...common,
      embassy: "Embassy / Consulate Visa Section",
      visaType: client.visaType,
      purpose: "Tourism and family leisure travel",
      travelDates: "03 Apr 2026 - 12 Apr 2026",
      funding: "Self-funded from savings and monthly employment income",
      background: "The applicant is employed full-time in Singapore and maintains continuing family and professional ties.",
      paragraphs:
        "I respectfully submit this explanation in support of the applicant's visa file. The enclosed supporting drafts summarize the applicant's travel purpose, financial evidence, accommodation arrangements, and employment background for consultant review before final submission.",
      closing: "Kindly review the attached supporting documents. Original documents should be checked before submission.",
    },
  };
  return fields[template];
};

const createDoc = (template: TemplateKey, client = demoClient): DocumentDraft => ({
  id: uid(),
  clientId: client.id,
  template,
  title: `${templateMeta[template].name} - ${client.name}`,
  status: "Draft",
  verified: false,
  watermark: true,
  updatedAt: new Date().toISOString(),
  fields: defaultFields(template, client),
  transactions: template === "financial" ? sampleTransactions : undefined,
  dayPlans: template === "itinerary" ? sampleDayPlans : undefined,
});

const createInitialStore = (): Store => ({
  clients: [demoClient],
  documents: (["financial", "employment", "salary", "itinerary", "accommodation", "cover"] as TemplateKey[]).map((template) => createDoc(template)),
  templates: Object.entries(templateMeta).map(([key, meta]) => ({ key: key as TemplateKey, name: meta.name })),
  transactions: sampleTransactions,
  activityLogs: [
    { id: uid(), text: "Demo client profile created", at: new Date().toISOString(), icon: "client" },
    { id: uid(), text: "Financial Evidence Summary sample ready for review", at: new Date().toISOString(), icon: "financial" },
  ],
  session: null,
});

const loadStore = (): Store => {
  const raw = localStorage.getItem("visahobe-document-studio");
  if (!raw) return createInitialStore();
  try {
    return { ...createInitialStore(), ...JSON.parse(raw) };
  } catch {
    return createInitialStore();
  }
};

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="space-y-2">
      <Label className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">{label}</Label>
      {children}
    </div>
  );
}

function Editable({ value, onChange, className }: { value: string | number | boolean | undefined; onChange: (value: string) => void; className?: string }) {
  return (
    <span
      className={cn("editable-field inline-block min-w-8 px-1", className)}
      contentEditable
      suppressContentEditableWarning
      onBlur={(event) => onChange(event.currentTarget.textContent || "")}
    >
      {String(value ?? "")}
    </span>
  );
}

function CountCard({ label, value, icon: Icon, tone }: { label: string; value: number; icon: typeof FileText; tone: string }) {
  return (
    <div className="studio-card overflow-hidden p-4 sm:p-5">
      <div className="flex items-start justify-between gap-3">
        <div>
          <p className="text-sm text-muted-foreground">{label}</p>
          <p className="mt-2 animate-count-pop text-3xl font-bold text-foreground">{value}</p>
        </div>
        <div className={cn("rounded-xl p-3 text-primary-foreground shadow-card", tone)}>
          <Icon className="h-5 w-5" />
        </div>
      </div>
    </div>
  );
}

function LoginScreen({ onLogin }: { onLogin: (session: NonNullable<Session>) => void }) {
  const [name, setName] = useState("VisaHOBe Consultant");
  const [role, setRole] = useState<Role>("Admin");
  return (
    <main className="min-h-screen bg-navy text-navy-foreground">
      <div className="mx-auto grid min-h-screen max-w-6xl items-center gap-8 px-4 py-10 lg:grid-cols-[1.05fr_0.95fr]">
        <section className="space-y-8">
          <div className="inline-flex items-center gap-3 rounded-full border border-primary-foreground/20 bg-primary-foreground/10 px-4 py-2 text-sm backdrop-blur">
            <ShieldCheck className="h-4 w-4 text-accent" /> Secure internal workspace
          </div>
          <div className="space-y-5">
            <p className="text-sm font-semibold uppercase tracking-[0.28em] text-secondary">VisaHOBe PTE. LTD.</p>
            <h1 className="max-w-3xl text-4xl font-bold leading-tight sm:text-6xl">VisaHOBe Document Studio</h1>
            <p className="max-w-2xl text-lg text-primary-foreground/75">
              Organize client profiles, prepare support-document drafts, verify originals, and produce print-ready internal visa file previews.
            </p>
          </div>
          <div className="grid gap-3 text-sm text-primary-foreground/80 sm:grid-cols-3">
            {[
              "Sample/Draft watermarking",
              "Original verification status",
              "A4 print preparation",
            ].map((item) => (
              <div key={item} className="rounded-xl border border-primary-foreground/15 bg-primary-foreground/10 p-4">
                <CheckCircle2 className="mb-3 h-5 w-5 text-accent" /> {item}
              </div>
            ))}
          </div>
        </section>
        <section className="studio-card border-primary-foreground/10 bg-card p-5 text-foreground sm:p-7">
          <div className="mb-6 flex items-center gap-3">
            <div className="rounded-xl bg-brand-gradient p-3 text-primary-foreground">
              <LogIn className="h-6 w-6" />
            </div>
            <div>
              <h2 className="text-xl font-bold">Mock admin sign in</h2>
              <p className="text-sm text-muted-foreground">No real authentication in this MVP. Session is stored locally.</p>
            </div>
          </div>
          <div className="space-y-4">
            <Field label="Consultant name">
              <Input value={name} onChange={(e) => setName(e.target.value)} />
            </Field>
            <Field label="Role">
              <Select value={role} onValueChange={(value) => setRole(value as Role)}>
                <SelectTrigger><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="Admin">Admin</SelectItem>
                  <SelectItem value="Documentation Officer">Documentation Officer</SelectItem>
                  <SelectItem value="Reviewer">Reviewer</SelectItem>
                </SelectContent>
              </Select>
            </Field>
            <Button variant="hero" className="w-full" onClick={() => onLogin({ name, role })}>
              <ShieldCheck className="h-4 w-4" /> Enter studio
            </Button>
          </div>
        </section>
      </div>
    </main>
  );
}

function ClientPanel({ clients, selectedClient, onSelect, onSave }: { clients: Client[]; selectedClient: Client; onSelect: (id: string) => void; onSave: (client: Client) => void }) {
  const [draft, setDraft] = useState<Client>(selectedClient);
  useEffect(() => setDraft(selectedClient), [selectedClient]);
  const update = (key: keyof Client, value: string) => setDraft((prev) => ({ ...prev, [key]: value }));
  return (
    <section className="grid gap-4 lg:grid-cols-[320px_1fr]">
      <div className="studio-card p-4">
        <div className="mb-4 flex items-center justify-between">
          <h2 className="font-bold">Client profiles</h2>
          <Button size="sm" variant="hero" onClick={() => {
            const client: Client = { ...demoClient, id: uid(), name: "New Client", passport: "", email: "", phone: "", notes: "", visaCountry: "", visaType: "" };
            onSave(client); onSelect(client.id); toast.success("New client profile added");
          }}><Plus className="h-4 w-4" /> Add</Button>
        </div>
        <div className="space-y-2">
          {clients.map((client) => (
            <button key={client.id} onClick={() => onSelect(client.id)} className={cn("w-full rounded-lg border p-3 text-left transition hover:bg-secondary/10", client.id === selectedClient.id ? "border-primary bg-secondary/10" : "border-border bg-card")}>
              <p className="font-semibold">{client.name}</p>
              <p className="text-xs text-muted-foreground">{client.visaCountry || "Country pending"} • {client.visaType || "Visa type pending"}</p>
            </button>
          ))}
        </div>
      </div>
      <div className="studio-card p-4 sm:p-5">
        <div className="mb-5 flex flex-wrap items-center justify-between gap-3">
          <div>
            <h2 className="text-xl font-bold">Client Profile Manager</h2>
            <p className="text-sm text-muted-foreground">Editable internal client details for document preparation.</p>
          </div>
          <Button variant="hero" onClick={() => onSave(draft)}><Save className="h-4 w-4" /> Save client</Button>
        </div>
        <div className="grid gap-4 md:grid-cols-2">
          {[
            ["Client name", "name"], ["Passport number", "passport"], ["Nationality", "nationality"], ["Date of birth", "dob"],
            ["Phone", "phone"], ["Email", "email"], ["Visa country", "visaCountry"], ["Visa type", "visaType"],
          ].map(([label, key]) => (
            <Field key={key} label={label}><Input type={key === "dob" ? "date" : "text"} value={String(draft[key as keyof Client] || "")} onChange={(e) => update(key as keyof Client, e.target.value)} /></Field>
          ))}
          <div className="md:col-span-2"><Field label="Address"><Textarea value={draft.address} onChange={(e) => update("address", e.target.value)} /></Field></div>
          <div className="md:col-span-2"><Field label="Notes"><Textarea value={draft.notes} onChange={(e) => update("notes", e.target.value)} /></Field></div>
        </div>
      </div>
    </section>
  );
}

function FieldEditor({ doc, onField, onTransactions, onPlans }: { doc: DocumentDraft; onField: (key: string, value: string | number | boolean) => void; onTransactions: (rows: TransactionRow[]) => void; onPlans: (rows: DayPlan[]) => void }) {
  const textField = (key: string, label: string, area = false) => (
    <Field label={label}>{area ? <Textarea value={String(doc.fields[key] || "")} onChange={(e) => onField(key, e.target.value)} /> : <Input value={String(doc.fields[key] || "")} onChange={(e) => onField(key, e.target.value)} />}</Field>
  );
  if (doc.template === "financial") return (
    <div className="space-y-4">
      <div className="grid gap-3 sm:grid-cols-2">
        {textField("clientName", "Client name")}{textField("accountHolderName", "Account holder")}{textField("accountNumber", "Account number")}{textField("bankName", "Bank name")}{textField("branch", "Branch")}{textField("statementPeriod", "Statement period")}{textField("currency", "Currency")}{textField("openingBalance", "Opening balance")}{textField("averageBalance", "Average balance")}
      </div>
      <label className="flex items-center gap-2 text-sm"><Checkbox checked={Boolean(doc.fields.maskAccount)} onCheckedChange={(v) => onField("maskAccount", Boolean(v))} /> Mask account number</label>
      <div className="rounded-xl border border-border p-3">
        <div className="mb-3 flex items-center justify-between"><h3 className="font-semibold">Transaction rows</h3><Button size="sm" variant="secondary" onClick={() => onTransactions([...(doc.transactions || []), { id: uid(), date: today(), description: "New row", reference: "", debit: 0, credit: 0, balance: 0 }])}><Plus className="h-4 w-4" /> Row</Button></div>
        <div className="space-y-3">
          {(doc.transactions || []).map((row) => <div key={row.id} className="grid gap-2 rounded-lg bg-surface p-2 sm:grid-cols-[1fr_1.6fr_1fr_0.8fr_0.8fr_0.8fr_auto]">
            {(["date", "description", "reference", "debit", "credit", "balance"] as const).map((key) => <Input key={key} type={["debit", "credit", "balance"].includes(key) ? "number" : key === "date" ? "date" : "text"} value={String(row[key])} onChange={(e) => onTransactions((doc.transactions || []).map((r) => r.id === row.id ? { ...r, [key]: ["debit", "credit", "balance"].includes(key) ? Number(e.target.value) : e.target.value } : r))} />)}
            <Button size="icon" variant="ghost" onClick={() => onTransactions((doc.transactions || []).filter((r) => r.id !== row.id))}><Trash2 className="h-4 w-4" /></Button>
          </div>)}
        </div>
      </div>
    </div>
  );
  if (doc.template === "itinerary") return (
    <div className="space-y-4">
      <div className="grid gap-3 sm:grid-cols-2">{textField("applicantName", "Applicant name")}{textField("country", "Country")}{textField("travelDates", "Travel dates")}{textField("flightRoute", "Flight route")}{textField("hotelCity", "Hotel / city")}{textField("purpose", "Purpose")}</div>
      <div className="rounded-xl border border-border p-3">
        <div className="mb-3 flex items-center justify-between"><h3 className="font-semibold">Daily plan</h3><Button size="sm" variant="secondary" onClick={() => onPlans([...(doc.dayPlans || []), { id: uid(), date: today(), city: "City", plan: "Plan details" }])}><Plus className="h-4 w-4" /> Day</Button></div>
        <div className="space-y-2">{(doc.dayPlans || []).map((row) => <div key={row.id} className="grid gap-2 rounded-lg bg-surface p-2 sm:grid-cols-[1fr_1fr_2fr_auto]"><Input type="date" value={row.date} onChange={(e) => onPlans((doc.dayPlans || []).map((r) => r.id === row.id ? { ...r, date: e.target.value } : r))} /><Input value={row.city} onChange={(e) => onPlans((doc.dayPlans || []).map((r) => r.id === row.id ? { ...r, city: e.target.value } : r))} /><Input value={row.plan} onChange={(e) => onPlans((doc.dayPlans || []).map((r) => r.id === row.id ? { ...r, plan: e.target.value } : r))} /><Button size="icon" variant="ghost" onClick={() => onPlans((doc.dayPlans || []).filter((r) => r.id !== row.id))}><Trash2 className="h-4 w-4" /></Button></div>)}</div>
      </div>
    </div>
  );
  const schemas: Record<Exclude<TemplateKey, "financial" | "itinerary">, [string, string, boolean?][]> = {
    employment: [["companyName", "Company name"], ["employeeName", "Employee name"], ["designation", "Designation"], ["joiningDate", "Joining date"], ["salary", "Salary"], ["hrPerson", "HR/contact person"], ["companyAddress", "Company address", true], ["issueDate", "Issue date"], ["bodyText", "Body text", true]],
    salary: [["employeeName", "Employee name"], ["designation", "Designation"], ["monthlySalary", "Monthly salary"], ["allowance", "Allowance"], ["netSalary", "Net salary"], ["companyName", "Company name"], ["issueDate", "Issue date"], ["authorizedPerson", "Authorized person"]],
    accommodation: [["applicantName", "Applicant name"], ["hotelName", "Hotel / host name"], ["address", "Address", true], ["checkIn", "Check-in"], ["checkOut", "Check-out"], ["bookingReference", "Booking reference"], ["contact", "Contact"], ["notes", "Notes", true]],
    cover: [["applicantName", "Applicant name"], ["embassy", "Embassy / consulate"], ["visaType", "Visa type"], ["purpose", "Purpose"], ["travelDates", "Travel dates"], ["funding", "Sponsor / funding"], ["background", "Employment/business background", true], ["paragraphs", "Explanation paragraphs", true], ["closing", "Closing", true]],
  };
  return <div className="grid gap-3 sm:grid-cols-2">{schemas[doc.template].map(([key, label, area]) => <div key={key} className={area ? "sm:col-span-2" : ""}>{textField(key, label, area)}</div>)}</div>;
}

function A4Preview({ doc, client, onField, onTransactions, onPlans }: { doc: DocumentDraft; client?: Client; onField: (key: string, value: string) => void; onTransactions: (rows: TransactionRow[]) => void; onPlans: (rows: DayPlan[]) => void }) {
  const f = doc.fields;
  const currency = String(f.currency || "SGD");
  const edit = (key: string, className?: string) => <Editable value={f[key]} onChange={(v) => onField(key, v)} className={className} />;
  const header = <><div className="flex items-start justify-between border-b border-border pb-5"><div><p className="text-xs font-semibold uppercase tracking-[0.28em] text-primary">VisaHOBe PTE. LTD.</p><h2 className="mt-2 text-2xl font-bold text-navy">{templateMeta[doc.template].name}</h2></div><Badge className={statusClass[doc.status]} variant="outline">{doc.verified ? "Verified" : doc.status}</Badge></div>{doc.watermark && <div aria-hidden className="pointer-events-none absolute inset-0 z-10 flex rotate-[-28deg] items-center justify-center text-7xl font-extrabold uppercase tracking-widest text-muted-foreground/15 print:text-foreground/10">SAMPLE / DRAFT</div>}</>;
  return (
    <article className="a4-page mx-auto overflow-hidden p-8 font-body text-sm leading-relaxed">
      {header}
      <div className="mt-6 space-y-6">
        {doc.template === "financial" && <>
          <div className="grid grid-cols-2 gap-3 rounded-xl bg-surface p-4 text-xs"><p><b>Client:</b> {edit("clientName")}</p><p><b>Account holder:</b> {edit("accountHolderName")}</p><p><b>Account no:</b> {edit("accountNumber")}</p><p><b>Bank:</b> {edit("bankName")}</p><p><b>Branch:</b> {edit("branch")}</p><p><b>Period:</b> {edit("statementPeriod")}</p></div>
          <div className="grid grid-cols-4 gap-2 text-center text-xs"><div className="rounded-lg border p-3"><p className="text-muted-foreground">Opening</p><b>{fmtMoney(f.openingBalance as number, currency)}</b></div><div className="rounded-lg border p-3"><p className="text-muted-foreground">Total debit</p><b>{fmtMoney(f.totalDebit as number, currency)}</b></div><div className="rounded-lg border p-3"><p className="text-muted-foreground">Total credit</p><b>{fmtMoney(f.totalCredit as number, currency)}</b></div><div className="rounded-lg border p-3"><p className="text-muted-foreground">Closing</p><b>{fmtMoney(f.closingBalance as number, currency)}</b></div></div>
          <table className="w-full border-collapse text-[10px]"><thead className="bg-surface"><tr>{["Date", "Description", "Ref", "Debit", "Credit", "Balance"].map(h => <th key={h} className="border p-2 text-left">{h}</th>)}</tr></thead><tbody>{(doc.transactions || []).map(row => <tr key={row.id}><td className="border p-2">{row.date}</td><td className="border p-2"><Editable value={row.description} onChange={(v) => onTransactions((doc.transactions || []).map(r => r.id === row.id ? { ...r, description: v } : r))} /></td><td className="border p-2">{row.reference}</td><td className="border p-2 text-right">{fmtMoney(row.debit, currency)}</td><td className="border p-2 text-right">{fmtMoney(row.credit, currency)}</td><td className="border p-2 text-right">{fmtMoney(row.balance, currency)}</td></tr>)}</tbody></table>
          <p className="border-t border-border pt-4 text-[11px] text-muted-foreground">Prepared for internal visa file review. Verify against original documents before submission.</p>
        </>}
        {doc.template === "employment" && <LetterBlock title="To Whom It May Concern"><p>This is to certify that {edit("employeeName")} is employed by {edit("companyName")} as {edit("designation")} since {edit("joiningDate")}. Current salary: {edit("salary")}.</p><p>{edit("bodyText", "min-w-full")}</p><p>Issued by {edit("hrPerson")}<br />{edit("companyAddress")}</p></LetterBlock>}
        {doc.template === "salary" && <LetterBlock title="Salary Certificate"><p>This certifies that {edit("employeeName")} is working as {edit("designation")} at {edit("companyName")}.</p><div className="rounded-xl border p-4"><p>Monthly salary: {edit("monthlySalary")}</p><p>Allowance: {edit("allowance")}</p><p>Net salary: {edit("netSalary")}</p></div><p>Authorized by: {edit("authorizedPerson")} on {edit("issueDate")}</p></LetterBlock>}
        {doc.template === "itinerary" && <><div className="grid grid-cols-2 gap-3 rounded-xl bg-surface p-4 text-xs"><p><b>Applicant:</b> {edit("applicantName")}</p><p><b>Country:</b> {edit("country")}</p><p><b>Dates:</b> {edit("travelDates")}</p><p><b>Route:</b> {edit("flightRoute")}</p><p><b>Hotel / city:</b> {edit("hotelCity")}</p><p><b>Purpose:</b> {edit("purpose")}</p></div><table className="w-full text-xs"><thead className="bg-surface"><tr><th className="border p-2 text-left">Date</th><th className="border p-2 text-left">City</th><th className="border p-2 text-left">Plan</th></tr></thead><tbody>{(doc.dayPlans || []).map(row => <tr key={row.id}><td className="border p-2">{row.date}</td><td className="border p-2"><Editable value={row.city} onChange={(v) => onPlans((doc.dayPlans || []).map(r => r.id === row.id ? { ...r, city: v } : r))} /></td><td className="border p-2"><Editable value={row.plan} onChange={(v) => onPlans((doc.dayPlans || []).map(r => r.id === row.id ? { ...r, plan: v } : r))} /></td></tr>)}</tbody></table></>}
        {doc.template === "accommodation" && <LetterBlock title="Accommodation Summary"><p>Applicant {edit("applicantName")} has accommodation details recorded as follows for internal file preparation.</p><div className="grid grid-cols-2 gap-3 rounded-xl border p-4"><p>Hotel/Host: {edit("hotelName")}</p><p>Reference: {edit("bookingReference")}</p><p>Check-in: {edit("checkIn")}</p><p>Check-out: {edit("checkOut")}</p><p>Contact: {edit("contact")}</p><p>Address: {edit("address")}</p></div><p>{edit("notes")}</p></LetterBlock>}
        {doc.template === "cover" && <LetterBlock title="Visa Explanation Letter"><p>To: {edit("embassy")}</p><p>Applicant: {edit("applicantName")}<br />Visa type: {edit("visaType")}<br />Travel dates: {edit("travelDates")}</p><p>{edit("paragraphs", "min-w-full")}</p><p>Purpose: {edit("purpose")}. Funding: {edit("funding")}.</p><p>{edit("background", "min-w-full")}</p><p>{edit("closing", "min-w-full")}</p></LetterBlock>}
      </div>
      <footer className="absolute bottom-6 left-8 right-8 space-y-2 border-t border-border pt-3 text-[10px] text-muted-foreground"><p className="font-semibold text-foreground">Prepared for internal visa file review. Verify against original documents before submission. This is not an official bank or government document.</p><div className="flex justify-between"><span>Client file: {client?.name || "Unassigned"}</span><span>Draft prepared in VisaHOBe Document Studio • {new Date(doc.updatedAt).toLocaleString()}</span></div></footer>
    </article>
  );
}

function LetterBlock({ title, children }: { title: string; children: React.ReactNode }) {
  return <div className="space-y-5"><h3 className="text-lg font-bold text-navy">{title}</h3><div className="space-y-4 text-sm">{children}</div><div className="pt-8"><p className="font-semibold">Authorized signature</p><div className="mt-8 w-56 border-t border-foreground" /></div></div>;
}

type ImportValidation = { rowIndex: number; field: string; message: string };

function validateTransactions(rows: TransactionRow[], openingBalance: number): { errors: ImportValidation[]; computed: { totalDebit: number; totalCredit: number } } {
  const errors: ImportValidation[] = [];
  const dateRe = /^\d{4}-\d{2}-\d{2}$/;
  let running = openingBalance;
  let totalDebit = 0;
  let totalCredit = 0;
  rows.forEach((r, i) => {
    if (!r.date || !dateRe.test(r.date)) errors.push({ rowIndex: i, field: "date", message: "Date must be YYYY-MM-DD" });
    if (!r.description?.trim()) errors.push({ rowIndex: i, field: "description", message: "Description is required" });
    if (Number.isNaN(Number(r.debit)) || Number(r.debit) < 0) errors.push({ rowIndex: i, field: "debit", message: "Debit must be ≥ 0" });
    if (Number.isNaN(Number(r.credit)) || Number(r.credit) < 0) errors.push({ rowIndex: i, field: "credit", message: "Credit must be ≥ 0" });
    if (Number(r.debit) > 0 && Number(r.credit) > 0) errors.push({ rowIndex: i, field: "credit", message: "Row cannot have both debit and credit" });
    totalDebit += Number(r.debit || 0);
    totalCredit += Number(r.credit || 0);
    running = running + Number(r.credit || 0) - Number(r.debit || 0);
    if (r.balance !== undefined && r.balance !== null && !Number.isNaN(Number(r.balance))) {
      if (Math.abs(Number(r.balance) - running) > 0.01) {
        errors.push({ rowIndex: i, field: "balance", message: `Running balance mismatch (expected ${running.toFixed(2)})` });
      }
    }
  });
  return { errors, computed: { totalDebit, totalCredit } };
}

function ImportTransactionsDialog({ open, onOpenChange, openingBalance, onImport }: { open: boolean; onOpenChange: (o: boolean) => void; openingBalance: number; onImport: (rows: TransactionRow[]) => void }) {
  const [rawRows, setRawRows] = useState<Record<string, string>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [map, setMap] = useState<{ date: string; description: string; debit: string; credit: string; reference: string; balance: string }>({ date: "", description: "", debit: "", credit: "", reference: "", balance: "" });
  const [errors, setErrors] = useState<ImportValidation[]>([]);
  const [preview, setPreview] = useState<TransactionRow[]>([]);
  const inputRef = useRef<HTMLInputElement>(null);

  const reset = () => { setRawRows([]); setHeaders([]); setErrors([]); setPreview([]); setMap({ date: "", description: "", debit: "", credit: "", reference: "", balance: "" }); };

  const handleFile = async (file: File) => {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "", raw: false });
      if (!json.length) { toast.error("No rows found in file"); return; }
      const hdrs = Object.keys(json[0]);
      setHeaders(hdrs);
      setRawRows(json.map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, String(v ?? "")]))));
      const guess = (terms: string[]) => hdrs.find((h) => terms.some((t) => h.toLowerCase().includes(t))) || "";
      setMap({
        date: guess(["date"]),
        description: guess(["desc", "narration", "particular", "details"]),
        debit: guess(["debit", "withdraw", "out"]),
        credit: guess(["credit", "deposit", "in"]),
        reference: guess(["ref", "txn", "transaction id"]),
        balance: guess(["balance"]),
      });
      toast.success(`Loaded ${json.length} rows. Confirm column mapping.`);
    } catch (err) {
      toast.error("Could not read file. Use a valid CSV or XLSX.");
    }
  };

  const buildRows = (): TransactionRow[] => {
    const num = (v: string) => Number((v || "").toString().replace(/[, ]/g, "")) || 0;
    return rawRows.map((r) => {
      const debit = num(r[map.debit] || "");
      const credit = num(r[map.credit] || "");
      const balance = map.balance ? num(r[map.balance] || "") : NaN;
      const rawDate = String(r[map.date] || "").trim();
      const isoMatch = rawDate.match(/^(\d{4})-(\d{2})-(\d{2})/);
      const dmYMatch = rawDate.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
      const date = isoMatch ? `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}` : dmYMatch ? `${dmYMatch[3]}-${dmYMatch[2].padStart(2, "0")}-${dmYMatch[1].padStart(2, "0")}` : rawDate.slice(0, 10);
      return { id: uid(), date, description: String(r[map.description] || ""), reference: String(r[map.reference] || ""), debit, credit, balance: Number.isNaN(balance) ? 0 : balance };
    });
  };

  const validate = () => {
    if (!map.date || !map.description) { toast.error("Map at least Date and Description columns"); return; }
    const rows = buildRows();
    const { errors: errs } = validateTransactions(rows, Number(openingBalance) || 0);
    setPreview(rows); setErrors(errs);
    if (errs.length) toast.error(`${errs.length} validation issue(s) found`); else toast.success("Validation passed — ready to import");
  };

  const handleImport = () => {
    if (!preview.length) { validate(); return; }
    if (errors.length) { toast.error("Fix validation errors before importing"); return; }
    onImport(preview);
    onOpenChange(false);
    reset();
  };

  return (
    <Dialog open={open} onOpenChange={(o) => { onOpenChange(o); if (!o) reset(); }}>
      <DialogContent className="max-w-2xl max-h-[90vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle>Import transactions (CSV / XLSX)</DialogTitle>
          <DialogDescription>Upload a bank export, then map columns. Validation runs before saving.</DialogDescription>
        </DialogHeader>
        <div className="space-y-4">
          <div className="rounded-xl border border-dashed border-border p-4 text-center">
            <Upload className="mx-auto mb-2 h-6 w-6 text-muted-foreground" />
            <input ref={inputRef} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])} />
            <Button variant="hero" onClick={() => inputRef.current?.click()}><Upload className="h-4 w-4" /> Choose file</Button>
            <p className="mt-2 text-xs text-muted-foreground">Supports .csv, .xlsx, .xls. First sheet is read.</p>
          </div>
          {headers.length > 0 && (
            <div className="grid gap-3 sm:grid-cols-2">
              {(["date", "description", "debit", "credit", "reference", "balance"] as const).map((key) => (
                <Field key={key} label={`${key[0].toUpperCase()}${key.slice(1)} column${key === "date" || key === "description" ? " *" : ""}`}>
                  <Select value={map[key] || "__none"} onValueChange={(v) => setMap((m) => ({ ...m, [key]: v === "__none" ? "" : v }))}>
                    <SelectTrigger><SelectValue placeholder="Select column" /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__none">— None —</SelectItem>
                      {headers.map((h) => <SelectItem key={h} value={h}>{h}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </Field>
              ))}
            </div>
          )}
          {rawRows.length > 0 && <div className="flex items-center justify-between text-xs text-muted-foreground"><span>{rawRows.length} rows ready.</span><Button size="sm" variant="outline" onClick={validate}>Validate rows</Button></div>}
          {errors.length > 0 && (
            <div className="rounded-lg border border-destructive/30 bg-destructive/5 p-3">
              <p className="mb-2 text-xs font-semibold text-destructive">{errors.length} validation issue(s):</p>
              <ul className="max-h-40 space-y-1 overflow-auto text-xs text-destructive">
                {errors.slice(0, 25).map((e, i) => <li key={i}>Row {e.rowIndex + 1} • <b>{e.field}</b>: {e.message}</li>)}
                {errors.length > 25 && <li>…and {errors.length - 25} more</li>}
              </ul>
            </div>
          )}
          {preview.length > 0 && !errors.length && <p className="rounded-lg border border-success/30 bg-success/5 p-2 text-xs text-success">All {preview.length} rows passed validation.</p>}
        </div>
        <DialogFooter>
          <Button variant="outline" onClick={() => onOpenChange(false)}>Cancel</Button>
          <Button variant="hero" disabled={!rawRows.length || (preview.length > 0 && errors.length > 0)} onClick={handleImport}>{preview.length ? "Import" : "Validate & continue"}</Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function ExportDialog({ open, onOpenChange, onConfirm }: { open: boolean; onOpenChange: (o: boolean) => void; onConfirm?: (mode: "print" | "pdf") => void }) {
  const [mode, setMode] = useState<"print" | "pdf">("pdf");
  const trigger = () => {
    onOpenChange(false);
    onConfirm?.(mode);
    toast.success(mode === "pdf" ? "Use 'Save as PDF' in the print dialog" : "Print dialog opened");
    setTimeout(() => window.print(), 150);
  };
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent>
        <DialogHeader>
          <DialogTitle>Export document</DialogTitle>
          <DialogDescription>A4 print styles are applied automatically. The browser's print dialog is used for PDF export.</DialogDescription>
        </DialogHeader>
        <div className="grid gap-3 sm:grid-cols-2">
          <button onClick={() => setMode("pdf")} className={cn("rounded-xl border p-4 text-left transition", mode === "pdf" ? "border-primary bg-secondary/10" : "border-border")}>
            <Download className="mb-2 h-5 w-5 text-primary" />
            <p className="font-semibold">Save as PDF</p>
            <p className="text-xs text-muted-foreground">Choose 'Save as PDF' as the destination in the print dialog.</p>
          </button>
          <button onClick={() => setMode("print")} className={cn("rounded-xl border p-4 text-left transition", mode === "print" ? "border-primary bg-secondary/10" : "border-border")}>
            <Printer className="mb-2 h-5 w-5 text-primary" />
            <p className="font-semibold">Print only</p>
            <p className="text-xs text-muted-foreground">Send directly to a connected A4 printer.</p>
          </button>
        </div>
        <DialogFooter>
          <Button variant="outline" onClick={() => onOpenChange(false)}>Cancel</Button>
          <Button variant="hero" onClick={trigger}>{mode === "pdf" ? "Open Save as PDF" : "Open print dialog"}</Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function GlobalSearchDialog({ open, onOpenChange, clients, documents, onPickDoc, onPickClient }: { open: boolean; onOpenChange: (o: boolean) => void; clients: Client[]; documents: DocumentDraft[]; onPickDoc: (id: string) => void; onPickClient: (id: string) => void }) {
  const [q, setQ] = useState("");
  useEffect(() => { if (!open) setQ(""); }, [open]);
  const term = q.trim().toLowerCase();
  const matchedClients = !term ? clients.slice(0, 5) : clients.filter((c) => `${c.name} ${c.passport} ${c.email} ${c.visaCountry} ${c.visaType}`.toLowerCase().includes(term));
  const matchedDocs = !term ? documents.slice(0, 8) : documents.filter((d) => {
    const c = clients.find((x) => x.id === d.clientId);
    return `${d.title} ${templateMeta[d.template].name} ${d.status} ${c?.name || ""} ${c?.passport || ""}`.toLowerCase().includes(term);
  });
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-xl">
        <DialogHeader>
          <DialogTitle>Search clients & documents</DialogTitle>
          <DialogDescription>Find any client profile or document draft across all templates.</DialogDescription>
        </DialogHeader>
        <div className="relative">
          <Search className="absolute left-3 top-2.5 h-4 w-4 text-muted-foreground" />
          <Input autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search by name, passport, title, status…" className="pl-9" />
        </div>
        <div className="max-h-80 space-y-4 overflow-auto">
          <div>
            <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground">Clients ({matchedClients.length})</p>
            <div className="space-y-1">
              {matchedClients.map((c) => (
                <button key={c.id} onClick={() => { onPickClient(c.id); onOpenChange(false); }} className="flex w-full items-center justify-between rounded-lg border border-border p-2 text-left text-sm hover:bg-secondary/10">
                  <span><b>{c.name}</b> <span className="text-muted-foreground">• {c.passport || "no passport"}</span></span>
                  <span className="text-xs text-muted-foreground">{c.visaCountry || "—"}</span>
                </button>
              ))}
              {!matchedClients.length && <p className="text-xs text-muted-foreground">No clients matched.</p>}
            </div>
          </div>
          <div>
            <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground">Documents ({matchedDocs.length})</p>
            <div className="space-y-1">
              {matchedDocs.map((d) => {
                const c = clients.find((x) => x.id === d.clientId);
                return (
                  <button key={d.id} onClick={() => { onPickDoc(d.id); onOpenChange(false); }} className="flex w-full items-center justify-between rounded-lg border border-border p-2 text-left text-sm hover:bg-secondary/10">
                    <span><b>{d.title}</b> <span className="text-muted-foreground">• {templateMeta[d.template].short}</span></span>
                    <span className="flex items-center gap-2 text-xs text-muted-foreground">{c?.name || "—"} <Badge className={statusClass[d.status]} variant="outline">{d.status}</Badge></span>
                  </button>
                );
              })}
              {!matchedDocs.length && <p className="text-xs text-muted-foreground">No documents matched.</p>}
            </div>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}

function DocumentStudio({ doc, clients, role, userName, onSave, onDuplicate, onReset, onLog }: { doc: DocumentDraft; clients: Client[]; role: Role; userName: string; onSave: (doc: DocumentDraft, message?: string) => void; onDuplicate: (doc: DocumentDraft) => void; onReset: () => void; onLog: (text: string, icon: ActivityLog["icon"]) => void }) {
  const selectedClient = clients.find((c) => c.id === doc.clientId) || clients[0];
  const [importOpen, setImportOpen] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [noteDraft, setNoteDraft] = useState("");
  const allowedStatuses = ROLE_PERMS[role].setStatus;
  const canVerify = can(role, "verify");
  const canSubmit = can(role, "submit");

  const setField = (key: string, value: string | number | boolean) => onSave({ ...doc, fields: { ...doc.fields, [key]: value }, updatedAt: new Date().toISOString() }, "Draft updated");
  const recalcFromRows = (rows: TransactionRow[]) => {
    const totalDebit = rows.reduce((s, r) => s + Number(r.debit || 0), 0);
    const totalCredit = rows.reduce((s, r) => s + Number(r.credit || 0), 0);
    const closingBalance = rows.length ? Number(rows[rows.length - 1].balance || 0) : Number(doc.fields.openingBalance || 0) + totalCredit - totalDebit;
    const averageBalance = rows.length ? Math.round(rows.reduce((s, r) => s + Number(r.balance || 0), 0) / rows.length) : 0;
    return { totalDebit, totalCredit, closingBalance, averageBalance };
  };
  const setTransactions = (rows: TransactionRow[], message = "Transactions recalculated") => onSave({ ...doc, transactions: rows, fields: { ...doc.fields, ...recalcFromRows(rows) }, updatedAt: new Date().toISOString() }, message);
  const setPlans = (rows: DayPlan[]) => onSave({ ...doc, dayPlans: rows, updatedAt: new Date().toISOString() }, "Itinerary updated");

  const updateStatus = (next: Status) => {
    if (!allowedStatuses.includes(next)) { toast.error(`Your role (${role}) cannot set status to ${next}`); return; }
    if (next === "Submitted" && !canSubmit) { toast.error("Only Admin or Documentation Officer can submit"); return; }
    if (next === "Verified" && !canVerify) { toast.error("Only Admin or Reviewer can verify"); return; }
    const verified = next === "Verified" ? true : doc.verified;
    onSave({ ...doc, status: next, verified, updatedAt: new Date().toISOString() }, `Status: ${next}`);
    onLog(`${userName} (${role}) set "${doc.title}" → ${next}`, "status");
  };

  const addNote = () => {
    const text = noteDraft.trim();
    if (!text) return;
    const note: ReviewNote = { id: uid(), author: userName, role, text, at: new Date().toISOString() };
    onSave({ ...doc, reviewNotes: [note, ...(doc.reviewNotes || [])], updatedAt: new Date().toISOString() }, "Review note added");
    onLog(`${userName} (${role}) added note on "${doc.title}": ${text.slice(0, 80)}`, "note");
    setNoteDraft("");
  };

  return (
    <section className="grid gap-4 xl:grid-cols-[440px_1fr]">
      <aside className="no-print studio-card h-fit p-4">
        <div className="mb-4 flex items-start justify-between gap-3">
          <div>
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Document Workspace</p>
            <h2 className="text-xl font-bold">{templateMeta[doc.template].name}</h2>
          </div>
          <Badge className={statusClass[doc.status]} variant="outline">{doc.status}</Badge>
        </div>
        <div className="mb-4 grid gap-3 sm:grid-cols-2">
          <Field label="Client">
            <Select value={doc.clientId} onValueChange={(v) => onSave({ ...doc, clientId: v, updatedAt: new Date().toISOString() }, "Client linked")}>
              <SelectTrigger><SelectValue /></SelectTrigger>
              <SelectContent>{clients.map(c => <SelectItem key={c.id} value={c.id}>{c.name}</SelectItem>)}</SelectContent>
            </Select>
          </Field>
          <Field label={`Status (role: ${role})`}>
            <Select value={doc.status} onValueChange={(v) => updateStatus(v as Status)}>
              <SelectTrigger><SelectValue /></SelectTrigger>
              <SelectContent>
                {STATUS_FLOW.map(s => (
                  <SelectItem key={s} value={s} disabled={!allowedStatuses.includes(s)}>
                    {s}{!allowedStatuses.includes(s) ? " (restricted)" : ""}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </Field>
        </div>
        <div className="mb-4 flex flex-wrap gap-3 text-sm">
          <label className={cn("flex items-center gap-2", !canVerify && "opacity-50")}>
            <Checkbox disabled={!canVerify} checked={doc.watermark} onCheckedChange={(v) => onSave({ ...doc, watermark: Boolean(v) }, "Watermark toggled")} />
            SAMPLE / DRAFT watermark
          </label>
          <label className={cn("flex items-center gap-2", !canVerify && "opacity-50")}>
            <Checkbox disabled={!canVerify} checked={doc.verified} onCheckedChange={(v) => updateStatus(Boolean(v) ? "Verified" : "Review")} />
            Verified against original {!canVerify && <Lock className="h-3 w-3" />}
          </label>
        </div>

        {doc.template === "financial" && (
          <div className="mb-3">
            <Button variant="secondary" className="w-full" onClick={() => setImportOpen(true)}><Upload className="h-4 w-4" /> Import CSV / XLSX</Button>
          </div>
        )}

        <FieldEditor doc={doc} onField={setField} onTransactions={(rows) => setTransactions(rows)} onPlans={setPlans} />

        <div className="mt-5 rounded-xl border border-border p-3">
          <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground">Review notes</p>
          <div className="mb-2 flex gap-2">
            <Input value={noteDraft} onChange={(e) => setNoteDraft(e.target.value)} placeholder="Add a review note…" />
            <Button variant="hero" size="sm" onClick={addNote}><Plus className="h-4 w-4" /></Button>
          </div>
          <div className="max-h-40 space-y-2 overflow-auto">
            {(doc.reviewNotes || []).map((n) => (
              <div key={n.id} className="rounded-lg bg-surface p-2 text-xs">
                <p className="font-semibold">{n.author} <span className="text-muted-foreground">• {n.role}</span></p>
                <p>{n.text}</p>
                <p className="text-muted-foreground">{new Date(n.at).toLocaleString()}</p>
              </div>
            ))}
            {!(doc.reviewNotes || []).length && <p className="text-xs text-muted-foreground">No notes yet.</p>}
          </div>
        </div>

        <div className="mt-5 grid grid-cols-2 gap-2">
          <Button variant="hero" onClick={() => onSave({ ...doc, updatedAt: new Date().toISOString() }, "Draft saved")}><Save className="h-4 w-4" /> Save</Button>
          <Button variant="secondary" onClick={() => onDuplicate(doc)}><Copy className="h-4 w-4" /> Duplicate</Button>
          <Button variant="outline" onClick={onReset}><RefreshCcw className="h-4 w-4" /> Reset sample</Button>
          <Button variant="outline" onClick={() => setExportOpen(true)}><Printer className="h-4 w-4" /> Print / Export</Button>
          {canSubmit && doc.status !== "Submitted" && (
            <Button variant="warm" className="col-span-2" onClick={() => updateStatus("Submitted")}><Send className="h-4 w-4" /> Mark as Submitted</Button>
          )}
        </div>
      </aside>
      <div className="print-area overflow-auto rounded-xl bg-surface-strong p-3 sm:p-6">
        <A4Preview doc={doc} client={selectedClient} onField={(k, v) => setField(k, v)} onTransactions={(rows) => setTransactions(rows)} onPlans={setPlans} />
      </div>
      <ImportTransactionsDialog open={importOpen} onOpenChange={setImportOpen} openingBalance={Number(doc.fields.openingBalance) || 0} onImport={(rows) => { setTransactions(rows, `Imported ${rows.length} transactions`); onLog(`${userName} (${role}) imported ${rows.length} transactions into "${doc.title}"`, "import"); }} />
      <ExportDialog open={exportOpen} onOpenChange={setExportOpen} onConfirm={(mode) => onLog(`${userName} (${role}) ${mode === "pdf" ? "exported PDF" : "printed"} "${doc.title}"`, "export")} />
    </section>
  );
}

const navItems = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "clients", label: "Clients", icon: UserRound },
  { id: "documents", label: "Documents", icon: FileText },
  { id: "activity", label: "Activity", icon: Activity },
] as const;
type View = typeof navItems[number]["id"];

export default function Index() {
  const [store, setStore] = useState<Store>(() => loadStore());
  const [view, setView] = useState<View>("dashboard");
  const [selectedClientId, setSelectedClientId] = useState(store.clients[0]?.id || "");
  const [selectedDocId, setSelectedDocId] = useState(store.documents[0]?.id || "");
  const [query, setQuery] = useState("");
  const [country, setCountry] = useState("all");
  const [visaType, setVisaType] = useState("all");
  const [globalSearchOpen, setGlobalSearchOpen] = useState(false);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "k") { e.preventDefault(); setGlobalSearchOpen(true); }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  useEffect(() => localStorage.setItem("visahobe-document-studio", JSON.stringify(store)), [store]);

  const addLog = (text: string, icon: ActivityLog["icon"] = "save") => setStore((prev) => ({ ...prev, activityLogs: [{ id: uid(), text, at: new Date().toISOString(), icon }, ...prev.activityLogs].slice(0, 20) }));
  const login = (session: NonNullable<Session>) => { setStore((prev) => ({ ...prev, session })); toast.success(`Welcome, ${session.role}`); };
  if (!store.session) return <LoginScreen onLogin={login} />;

  const selectedClient = store.clients.find((c) => c.id === selectedClientId) || store.clients[0];
  const selectedDoc = store.documents.find((d) => d.id === selectedDocId) || store.documents[0];
  const countries = Array.from(new Set(store.clients.map((c) => c.visaCountry).filter(Boolean)));
  const visaTypes = Array.from(new Set(store.clients.map((c) => c.visaType).filter(Boolean)));
  const filteredDocs = store.documents.filter((doc) => {
    const c = store.clients.find((client) => client.id === doc.clientId);
    const hay = `${doc.title} ${templateMeta[doc.template].name} ${c?.name} ${c?.passport}`.toLowerCase();
    return hay.includes(query.toLowerCase()) && (country === "all" || c?.visaCountry === country) && (visaType === "all" || c?.visaType === visaType);
  });
  const pendingDocuments = store.documents.filter((d) => !d.verified || ["Draft", "Review"].includes(d.status)).length;
  const completedFiles = store.documents.filter((d) => ["Ready", "Submitted", "Verified"].includes(d.status)).length;

  const saveDoc = (doc: DocumentDraft, message = "Draft saved") => { setStore((prev) => ({ ...prev, documents: prev.documents.map((d) => d.id === doc.id ? doc : d) })); toast.success(message); };
  const saveClient = (client: Client) => { setStore((prev) => ({ ...prev, clients: prev.clients.some((c) => c.id === client.id) ? prev.clients.map((c) => c.id === client.id ? client : c) : [...prev.clients, client] })); addLog(`Client profile saved: ${client.name}`, "client"); toast.success("Client profile saved"); };
  const duplicateDoc = (doc: DocumentDraft) => { const copyDoc = { ...doc, id: uid(), title: `${doc.title} copy`, updatedAt: new Date().toISOString() }; setStore((prev) => ({ ...prev, documents: [copyDoc, ...prev.documents] })); setSelectedDocId(copyDoc.id); addLog(`Duplicated ${templateMeta[doc.template].short} draft`, doc.template); toast.success("Document duplicated"); };
  const resetDoc = () => { const client = store.clients.find((c) => c.id === selectedDoc.clientId) || demoClient; const fresh = { ...createDoc(selectedDoc.template, client), id: selectedDoc.id, clientId: selectedDoc.clientId, status: "Draft" as Status }; saveDoc(fresh, "Sample data restored"); };
  const createNewDoc = (template: TemplateKey) => { const doc = createDoc(template, selectedClient || demoClient); setStore((prev) => ({ ...prev, documents: [doc, ...prev.documents] })); setSelectedDocId(doc.id); setView("documents"); addLog(`Created ${templateMeta[template].short} draft`, template); toast.success("New document draft created"); };

  const sidebar = <aside className="no-print hidden w-72 shrink-0 border-r border-primary-foreground/10 bg-navy p-4 text-navy-foreground lg:block"><div className="mb-8 flex items-center gap-3"><div className="rounded-xl bg-hot-gradient p-3"><FileCheck2 className="h-6 w-6" /></div><div><p className="text-xs uppercase tracking-[0.22em] text-secondary">VisaHOBe</p><p className="font-bold">Document Studio</p></div></div><nav className="space-y-2">{navItems.map((item) => <Button key={item.id} variant="nav" className={cn("w-full text-primary-foreground/70", view === item.id && "bg-primary-foreground/10 text-primary-foreground")} onClick={() => setView(item.id)}><item.icon className="h-4 w-4" /> {item.label}</Button>)}</nav><div className="mt-8 rounded-xl border border-primary-foreground/15 bg-primary-foreground/10 p-4 text-sm"><p className="font-semibold">{store.session.name}</p><p className="text-primary-foreground/70">{store.session.role}</p><Button className="mt-4 w-full" variant="warm" size="sm" onClick={() => setStore((prev) => ({ ...prev, session: null }))}><LogOut className="h-4 w-4" /> Sign out</Button></div></aside>;

  return (
    <main className="min-h-screen bg-background">
      <div className="flex min-h-screen">
        {sidebar}
        <section className="min-w-0 flex-1 pb-24 lg:pb-0">
          <header className="no-print sticky top-0 z-20 border-b border-border bg-background/90 px-4 py-3 backdrop-blur lg:px-6">
            <div className="flex flex-col gap-3 xl:flex-row xl:items-center xl:justify-between">
              <div className="flex items-center justify-between gap-3"><div><p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">VisaHOBe PTE. LTD.</p><h1 className="text-xl font-bold">{view === "dashboard" ? "Operations Dashboard" : view === "clients" ? "Client Profile Manager" : view === "documents" ? "Document Workspace" : "Recent Activity"}</h1></div><Menu className="h-5 w-5 text-muted-foreground lg:hidden" /></div>
              <div className="grid gap-2 sm:grid-cols-[1.5fr_1fr_1fr_auto_auto] xl:w-[820px]"><button onClick={() => setGlobalSearchOpen(true)} className="relative flex items-center rounded-md border border-input bg-background pl-9 pr-3 text-left text-sm text-muted-foreground hover:bg-secondary/10"><Search className="absolute left-3 top-2.5 h-4 w-4" />Search clients & documents… <kbd className="ml-auto rounded bg-muted px-1.5 py-0.5 text-[10px]">⌘K</kbd></button><Select value={country} onValueChange={setCountry}><SelectTrigger><SelectValue placeholder="Country" /></SelectTrigger><SelectContent><SelectItem value="all">All countries</SelectItem>{countries.map(c => <SelectItem key={c} value={c}>{c}</SelectItem>)}</SelectContent></Select><Select value={visaType} onValueChange={setVisaType}><SelectTrigger><SelectValue placeholder="Visa type" /></SelectTrigger><SelectContent><SelectItem value="all">All visa types</SelectItem>{visaTypes.map(v => <SelectItem key={v} value={v}>{v}</SelectItem>)}</SelectContent></Select><Input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Filter list…" /><Button variant="hero" onClick={() => createNewDoc("financial")}><Plus className="h-4 w-4" /> Draft</Button></div>
            </div>
          </header>
          <div className="p-4 lg:p-6">
            {view === "dashboard" && <div className="space-y-6"><div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4"><CountCard label="Total clients" value={store.clients.length} icon={UserRound} tone="bg-brand-gradient" /><CountCard label="Drafts" value={store.documents.filter(d => d.status === "Draft").length} icon={FileText} tone="bg-secondary" /><CountCard label="Pending documents" value={pendingDocuments} icon={ClipboardList} tone="bg-hot-gradient" /><CountCard label="Completed files" value={completedFiles} icon={ShieldCheck} tone="bg-success" /></div><section className="grid gap-4 xl:grid-cols-[1fr_380px]"><div className="studio-card p-4"><div className="mb-4 flex items-center justify-between"><h2 className="font-bold">Quick actions</h2><Badge variant="outline">{filteredDocs.length} visible drafts</Badge></div><div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-3">{(Object.keys(templateMeta) as TemplateKey[]).map((key) => { const Icon = templateMeta[key].icon; return <button key={key} onClick={() => createNewDoc(key)} className="rounded-xl border border-border bg-card p-4 text-left transition hover:-translate-y-0.5 hover:border-primary hover:shadow-card"><Icon className="mb-3 h-5 w-5 text-primary" /><p className="font-semibold">{templateMeta[key].short}</p><p className="text-xs text-muted-foreground">Create editable draft</p></button>; })}</div></div><div className="studio-card p-4"><h2 className="mb-4 font-bold">Recent activity</h2><div className="space-y-3">{store.activityLogs.slice(0, 6).map((log) => <div key={log.id} className="flex gap-3 rounded-lg bg-surface p-3"><Activity className="h-4 w-4 text-primary" /><div><p className="text-sm font-medium">{log.text}</p><p className="text-xs text-muted-foreground">{new Date(log.at).toLocaleString()}</p></div></div>)}</div></div></section></div>}
            {view === "clients" && selectedClient && (can(store.session.role, "manageClients") ? <ClientPanel clients={store.clients} selectedClient={selectedClient} onSelect={setSelectedClientId} onSave={saveClient} /> : <div className="studio-card p-6 text-sm"><Lock className="mb-2 h-5 w-5 text-muted-foreground" /><p className="font-semibold">Client management is restricted</p><p className="text-muted-foreground">Your role ({store.session.role}) can review documents but cannot edit client profiles.</p></div>)}
            {view === "documents" && <div className="space-y-4"><div className="no-print studio-card flex flex-col gap-3 p-3 lg:flex-row lg:items-center"><div className="grid flex-1 gap-2 md:grid-cols-3">{filteredDocs.map((doc) => { const Icon = templateMeta[doc.template].icon; return <button key={doc.id} onClick={() => setSelectedDocId(doc.id)} className={cn("rounded-lg border p-3 text-left transition hover:bg-secondary/10", selectedDocId === doc.id ? "border-primary bg-secondary/10" : "border-border")}><div className="flex items-start gap-2"><Icon className="mt-0.5 h-4 w-4 text-primary" /><div className="min-w-0"><p className="truncate text-sm font-semibold">{doc.title}</p><p className="text-xs text-muted-foreground">{templateMeta[doc.template].short} • {doc.status}</p></div></div></button>; })}</div></div>{selectedDoc && <DocumentStudio doc={selectedDoc} clients={store.clients} role={store.session.role} userName={store.session.name} onSave={saveDoc} onDuplicate={duplicateDoc} onReset={resetDoc} onLog={addLog} />}</div>}
            {view === "activity" && <div className="studio-card p-4"><h2 className="mb-4 text-xl font-bold">Activity Logs</h2><div className="space-y-3">{store.activityLogs.map((log) => <div key={log.id} className="flex items-center justify-between rounded-xl border border-border p-4"><div className="flex items-center gap-3"><CalendarDays className="h-5 w-5 text-primary" /><div><p className="font-medium">{log.text}</p><p className="text-sm text-muted-foreground">{new Date(log.at).toLocaleString()}</p></div></div><Badge variant="outline">Internal</Badge></div>)}</div></div>}
          </div>
        </section>
      </div>
      <nav className="no-print fixed bottom-0 left-0 right-0 z-30 grid grid-cols-4 border-t border-border bg-card p-2 shadow-card lg:hidden">{navItems.map((item) => <button key={item.id} onClick={() => setView(item.id)} className={cn("flex flex-col items-center gap-1 rounded-lg px-2 py-2 text-xs text-muted-foreground", view === item.id && "bg-secondary/10 text-primary")}><item.icon className="h-5 w-5" />{item.label}</button>)}</nav>
      <GlobalSearchDialog open={globalSearchOpen} onOpenChange={setGlobalSearchOpen} clients={store.clients} documents={store.documents} onPickDoc={(id) => { setSelectedDocId(id); setView("documents"); }} onPickClient={(id) => { setSelectedClientId(id); setView("clients"); }} />
    </main>
  );
}
