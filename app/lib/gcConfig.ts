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
