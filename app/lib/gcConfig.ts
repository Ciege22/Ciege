export interface GcConfigEntry {
  gc: string
  cr_match: string[]
  spo_match: string[]
  cr_label: string
  spo_label: string
}

export const GC_CONFIG: GcConfigEntry[] = [
  { gc: 'Mastec',   cr_match: ['mastec'],           spo_match: ['mastec network solutions llc'],                                          cr_label: 'MasTec',    spo_label: 'MASTEC_NETWORK_SOLUTIONS_LLC' },
  { gc: 'MZI',      cr_match: ['mzi group'],         spo_match: ['mzi group inc'],                                                          cr_label: 'MZI_Group', spo_label: 'MZI_GROUP_INC' },
  { gc: 'NV Tel',   cr_match: ['nv-tel', 'nv tel'],  spo_match: ['nv tel inc'],                                                             cr_label: 'NV-Tel',    spo_label: 'NV_TEL_INC' },
  { gc: 'Tech CX',  cr_match: ['tech cx'],            spo_match: ['tech construction inc'],                                                  cr_label: 'Tech_Cx',   spo_label: 'TECH_CONSTRUCTION_INC' },
  { gc: 'Vikor',    cr_match: ['vikor'],              spo_match: ['sioux falls tower specialists inc.', 'sioux falls tower specialists inc'], cr_label: 'Vikor',     spo_label: 'SIOUX_FALLS_TOWER_SPECIALISTS_INC' },
  { gc: 'TCE',      cr_match: ['tce'],                spo_match: ['tower communications experts. llc', 'tower communications experts llc'],  cr_label: 'TCE',       spo_label: 'TOWER_COMMUNICATIONS_EXPERTS_LLC' },
  { gc: 'InSite',   cr_match: ['insite'],             spo_match: ['insite telecom llc'],                                                     cr_label: 'InSite',    spo_label: 'INSITE_TELECOM_LLC' },
  // Vendor names below confirmed directly against the source SPO Report /
  // CR Tracker files by CJ (2026-09-02) — not guessed.
  { gc: 'Capital Tower', cr_match: ['capital tower'],  spo_match: ['capital tower and communications'],   cr_label: 'Capital_Tower', spo_label: 'CAPITAL_TOWER_AND_COMMUNICATIONS' },
  { gc: 'Elite',         cr_match: ['elite'],          spo_match: ['elite microwave solutions'],          cr_label: 'Elite',         spo_label: 'ELITE_MICROWAVE_SOLUTIONS' },
  { gc: 'Ethos',         cr_match: ['ethos'],          spo_match: ['ethos distributed solutions inc'],    cr_label: 'Ethos',         spo_label: 'ETHOS_DISTRIBUTED_SOLUTIONS_INC' },
  { gc: 'Site Property', cr_match: ['site properties'],spo_match: ['site property company llc'],          cr_label: 'Site_Property', spo_label: 'SITE_PROPERTY_COMPANY_LLC' },
  { gc: 'Steimel',       cr_match: ['steimel communications inc'], spo_match: ['steimel communications inc'], cr_label: 'Steimel',   spo_label: 'STEIMEL_COMMUNICATIONS_INC' },
  { gc: 'Viking',        cr_match: ['viking'],         spo_match: ['viking maintenance ltd dba finish'],  cr_label: 'Viking',        spo_label: 'VIKING_MAINTENANCE_LTD_DBA_FINISH' },
  { gc: 'WaveLink',      cr_match: ['wavelink'],        spo_match: ['wave link corp. of puerto rico'],     cr_label: 'WaveLink',      spo_label: 'WAVE_LINK_CORP_OF_PUERTO_RICO' },
]

export const SPO_VENDOR_COL_IN_MASTER = 43
export const CR_SUPPLIER_COL_IN_MASTER = 1

export function matches(value: unknown, matchList: string[]): boolean {
  if (!value) return false
  const v = String(value).trim().toLowerCase()
  return matchList.some(m => v === m.toLowerCase())
}

export function gcForVendor(vendor: unknown): string {
  const entry = GC_CONFIG.find(cfg => matches(vendor, cfg.spo_match))
  return entry ? entry.gc : String(vendor || '').trim()
}
