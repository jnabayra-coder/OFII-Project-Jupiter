import * as XLSX from 'xlsx';
import { OFIIFieldKey } from '../types';

export interface OFIISharedColumnDefinition {
  excelHeader: string;
  systemField: OFIIFieldKey;
  label: string;
  required: boolean;
  category: 'Required' | 'Identification' | 'Cargo' | 'Delivery' | 'POD';
  dataType: 'string' | 'number' | 'date' | 'time';
  description: string;
  aliases: string[];
  modules: ('forwarding' | 'dispatch')[];
  columnWidth: number;
  sampleValueForwarding?: any;
  sampleValueDispatch?: any;
}

export interface OFIIFieldDefinition {
  key: OFIIFieldKey;
  label: string;
  required: boolean;
  category: 'Required' | 'Identification' | 'Cargo' | 'Delivery' | 'POD';
  description: string;
  aliases: string[];
}

/**
 * Single Unified Column Schema for both Template Generation and Importer Column Mapping.
 */
export const OFII_SHARED_SCHEMA: OFIISharedColumnDefinition[] = [
  // 1. Forwarding & Core Shared Columns
  {
    excelHeader: 'Month',
    systemField: 'month',
    label: 'Month',
    required: false,
    category: 'Identification',
    dataType: 'string',
    description: 'Operational tracking month (e.g. August 2026)',
    aliases: ['month', 'billing month', 'period', 'cycle', 'month started'],
    modules: ['forwarding'],
    columnWidth: 16,
    sampleValueForwarding: 'August 2026',
  },
  {
    excelHeader: 'Coordinator',
    systemField: 'coordinator',
    label: 'Coordinator',
    required: false,
    category: 'Identification',
    dataType: 'string',
    description: 'Account coordinator (Auto-assigned from Client Name if omitted)',
    aliases: ['coordinator', 'account coordinator', 'assigned coordinator', 'ops coordinator', 'coordinator name'],
    modules: ['forwarding'],
    columnWidth: 22,
    sampleValueForwarding: 'Rojaylene Baldesancho',
  },
  {
    excelHeader: 'Client Name',
    systemField: 'client',
    label: 'Client Name',
    required: true,
    category: 'Required',
    dataType: 'string',
    description: 'Registered client enterprise (e.g. Intelligent Skin Care Inc., Philippine Charity Sweepstakes Office, RefaMED)',
    aliases: ['client name', 'client', 'customer', 'customer name', 'account', 'shipper', 'principal', 'client / account', 'client_name'],
    modules: ['forwarding', 'dispatch'],
    columnWidth: 38,
    sampleValueForwarding: 'Intelligent Skin Care Inc.',
    sampleValueDispatch: 'Philippine Charity Sweepstakes Office',
  },
  {
    excelHeader: 'Mode of Shipment',
    systemField: 'modeOfShipment',
    label: 'Mode of Shipment',
    required: true,
    category: 'Required',
    dataType: 'string',
    description: 'Transport method: Air, Land, Sea, or RORO',
    aliases: ['mode of shipment', 'mode', 'shipment mode', 'transport mode', 'freight mode', 'service type', 'shipping mode', 'delivery type', 'delivery mode', 'freight type'],
    modules: ['forwarding'],
    columnWidth: 18,
    sampleValueForwarding: 'Air',
  },
  {
    excelHeader: 'Delivery Area',
    systemField: 'area',
    label: 'Delivery Area',
    required: true,
    category: 'Required',
    dataType: 'string',
    description: 'Philippine destination zone: Luzon, Visayas, or Mindanao',
    aliases: ['delivery area', 'area', 'region', 'destination area', 'zone', 'island group', 'philippine area'],
    modules: ['forwarding'],
    columnWidth: 16,
    sampleValueForwarding: 'Luzon',
  },
  {
    excelHeader: 'Actual Dispatched Date',
    systemField: 'actualDispatchDate',
    label: 'Actual Dispatched Date',
    required: true,
    category: 'Required',
    dataType: 'date',
    description: 'Date freight left origin hub (MM/DD/YYYY or YYYY-MM-DD)',
    aliases: ['actual dispatched date', 'actual dispatch date', 'dispatch date', 'dispatched date', 'departure date', 'date dispatched', 'ship date', 'etd', 'dispatching date'],
    modules: ['forwarding'],
    columnWidth: 22,
    sampleValueForwarding: '08/25/2026',
  },
  {
    excelHeader: 'Consignee',
    systemField: 'consignee',
    label: 'Consignee',
    required: true,
    category: 'Required',
    dataType: 'string',
    description: 'Recipient business or store name',
    aliases: ['consignee', 'receiver', 'recipient', 'deliver to', 'consignee name', 'destination store', 'store'],
    modules: ['forwarding', 'dispatch'],
    columnWidth: 36,
    sampleValueForwarding: 'Watsons Personal Care Stores - SM Megamall',
    sampleValueDispatch: 'Mercury Drug - Cebu Main',
  },
  {
    excelHeader: 'Destination Code',
    systemField: 'destinationCode',
    label: 'Destination Code',
    required: false,
    category: 'Cargo',
    dataType: 'string',
    description: 'Specific port or city code (e.g. MNL-01, CDO-01, CEB-01, DVO-01)',
    aliases: ['destination code', 'dest code', 'port code', 'city code', 'hub', 'destination city', 'terminal code'],
    modules: ['forwarding'],
    columnWidth: 18,
    sampleValueForwarding: 'MNL-01',
  },
  {
    excelHeader: 'Quantity',
    systemField: 'quantity',
    label: 'Quantity',
    required: false,
    category: 'Cargo',
    dataType: 'number',
    description: 'Number of boxes or packages',
    aliases: ['quantity', 'qty', 'cases', 'boxes', 'packages', 'pkgs', 'pieces', 'units', 'cartons', 'quantity cases boxes'],
    modules: ['forwarding', 'dispatch'],
    columnWidth: 12,
    sampleValueForwarding: 150,
    sampleValueDispatch: 120,
  },
  {
    excelHeader: 'CBM',
    systemField: 'cbm',
    label: 'CBM',
    required: false,
    category: 'Cargo',
    dataType: 'number',
    description: 'Cubic volume for Sea / Air Freight',
    aliases: ['cbm', 'volume cbm', 'cubic meter', 'cubic meters'],
    modules: ['forwarding'],
    columnWidth: 10,
    sampleValueForwarding: 4.5,
  },
  {
    excelHeader: 'Reference Number',
    systemField: 'referenceNumber',
    label: 'Reference Number',
    required: true,
    category: 'Required',
    dataType: 'string',
    description: 'Unique project code or shipment reference (e.g. PRJ-ISCI-901)',
    aliases: ['reference number', 'ref no', 'ref no.', 'ref number', 'reference no', 'reference no.', 'project code', 'ref', 'so number', 'sales order'],
    modules: ['forwarding'],
    columnWidth: 42,
    sampleValueForwarding: 'PRJ-ISCI-901 [SAMPLE — DELETE BEFORE IMPORT]',
  },
  {
    excelHeader: 'Volume Weight',
    systemField: 'volumeWeightKg',
    label: 'Volume Weight',
    required: false,
    category: 'Cargo',
    dataType: 'number',
    description: 'Volumetric weight for Air Freight in kilograms',
    aliases: ['volume weight', 'volumetric weight', 'vol weight', 'volume weight kg', 'volume weight (kg)'],
    modules: ['forwarding'],
    columnWidth: 16,
    sampleValueForwarding: 750,
  },
  {
    excelHeader: 'Actual Weight',
    systemField: 'actualWeightKg',
    label: 'Actual Weight',
    required: false,
    category: 'Cargo',
    dataType: 'number',
    description: 'Gross cargo weight in kilograms',
    aliases: ['actual weight', 'weight', 'weight kg', 'weight (kg)', 'gross weight', 'actual weight kg', 'total weight kg', 'actual weight (kg)'],
    modules: ['forwarding'],
    columnWidth: 16,
    sampleValueForwarding: 1200,
  },
  {
    excelHeader: 'Fees',
    systemField: 'chargeableWeightFees',
    label: 'Fees / Chargeable Weight',
    required: false,
    category: 'Cargo',
    dataType: 'string',
    description: 'Billed amount or chargeable rate (e.g. PHP 35,000.00)',
    aliases: ['fees', 'chargeable fees', 'freight fee', 'rate', 'freight charges', 'chargeable weight fees', 'billing amount', 'amount', 'fees chargeable weight'],
    modules: ['forwarding'],
    columnWidth: 18,
    sampleValueForwarding: 'PHP 35,000.00',
  },
  {
    excelHeader: 'RDD',
    systemField: 'requestDeliveryDate',
    label: 'Request Delivery Date (RDD)',
    required: false,
    category: 'Delivery',
    dataType: 'date',
    description: 'Client-specific requested delivery deadline (MM/DD/YYYY or YYYY-MM-DD)',
    aliases: ['rdd', 'request delivery date', 'requested delivery date', 'client requested delivery date', 'client rdd', 'required delivery date', 'client delivery requirement', 'client deadline'],
    modules: ['forwarding'],
    columnWidth: 14,
    sampleValueForwarding: '08/30/2026',
  },
  {
    excelHeader: 'Actual Delivery Date',
    systemField: 'actualDeliveryDate',
    label: 'Delivery Date',
    required: false,
    category: 'Delivery',
    dataType: 'date',
    description: 'Date package was received by consignee (MM/DD/YYYY or YYYY-MM-DD)',
    aliases: ['actual delivery date', 'delivery date', 'date delivered', 'delivered date', 'received date'],
    modules: ['forwarding'],
    columnWidth: 20,
    sampleValueForwarding: '08/29/2026',
  },
  {
    excelHeader: 'POD Number',
    systemField: 'podNumber',
    label: 'POD Number',
    required: false,
    category: 'Identification',
    dataType: 'string',
    description: 'Proof of Delivery document number (e.g. POD-94108)',
    aliases: ['pod number', 'pod', 'pod no', 'pod no.', 'dr number', 'dr no', 'delivery receipt'],
    modules: ['forwarding'],
    columnWidth: 16,
    sampleValueForwarding: 'POD-94108',
  },
  {
    excelHeader: 'Actual POD Return Date',
    systemField: 'dateOfPodReturn',
    label: 'POD Return Date',
    required: false,
    category: 'POD',
    dataType: 'date',
    description: 'Date signed hardcopy POD was returned to OFII hub',
    aliases: ['actual pod return date', 'pod return date', 'date of pod return', 'pod date', 'pod received date', 'date pod returned', 'actual pod return'],
    modules: ['forwarding'],
    columnWidth: 22,
    sampleValueForwarding: '09/01/2026',
  },

  // 2. Daily Dispatch Monitoring Columns
  {
    excelHeader: 'Date',
    systemField: 'actualDispatchDate',
    label: 'Dispatch Date',
    required: true,
    category: 'Required',
    dataType: 'date',
    description: 'Date of vehicle dispatch (MM/DD/YYYY or YYYY-MM-DD)',
    aliases: ['date', 'dispatch date', 'actual dispatch date', 'date dispatched', 'trip date'],
    modules: ['dispatch'],
    columnWidth: 15,
    sampleValueDispatch: '08/25/2026',
  },
  {
    excelHeader: 'POD',
    systemField: 'podNumber',
    label: 'POD Number',
    required: true,
    category: 'Identification',
    dataType: 'string',
    description: 'Proof of delivery / dispatch number (e.g. POD-88101)',
    aliases: ['pod', 'pod number', 'pod no', 'pod no.', 'dr number', 'dr no'],
    modules: ['dispatch'],
    columnWidth: 15,
    sampleValueDispatch: 'POD-88101',
  },
  {
    excelHeader: 'Delivery Type',
    systemField: 'modeOfShipment',
    label: 'Delivery Type / Mode',
    required: false,
    category: 'Required',
    dataType: 'string',
    description: 'Delivery classification or mode (e.g. Land Freight, Sea Freight, Air Freight)',
    aliases: ['delivery type', 'delivery mode', 'service type', 'freight mode'],
    modules: ['dispatch'],
    columnWidth: 18,
    sampleValueDispatch: 'Sea Freight',
  },
  {
    excelHeader: 'Place',
    systemField: 'destination',
    label: 'Place / Destination',
    required: false,
    category: 'Cargo',
    dataType: 'string',
    description: 'Drop-off location / city / terminal destination',
    aliases: ['place', 'destination', 'location', 'delivery destination', 'destination address'],
    modules: ['dispatch'],
    columnWidth: 32,
    sampleValueDispatch: 'Cebu Fuente Osmena Terminal',
  },
  {
    excelHeader: 'Truck Used',
    systemField: 'courier',
    label: 'Truck Used / Provider',
    required: false,
    category: 'Cargo',
    dataType: 'string',
    description: 'Type of vehicle or carrier provider',
    aliases: ['truck used', 'truck type', 'vehicle type', 'truck provider', 'carrier', 'hauler', 'courier'],
    modules: ['dispatch'],
    columnWidth: 20,
    sampleValueDispatch: '10W Wing Van',
  },
  {
    excelHeader: 'Plate Number',
    systemField: 'plateNumber',
    label: 'Plate Number',
    required: false,
    category: 'Cargo',
    dataType: 'string',
    description: 'Vehicle license plate number (e.g. NDB-4921)',
    aliases: ['plate number', 'truck plate', 'plate no', 'plate no.', 'vehicle plate'],
    modules: ['dispatch'],
    columnWidth: 15,
    sampleValueDispatch: 'NDB-4921',
  },
  {
    excelHeader: 'Time Arrived',
    systemField: 'timeArrived',
    label: 'Time Arrived',
    required: false,
    category: 'Delivery',
    dataType: 'time',
    description: 'Actual time vehicle arrived at compound (e.g. 07:30 AM)',
    aliases: ['time arrived', 'arrival time', 'truck arrival', 'truck arrival time', 'time in', 'gate in'],
    modules: ['dispatch'],
    columnWidth: 16,
    sampleValueDispatch: '07:30 AM',
  },
  {
    excelHeader: 'Start Loading Time',
    systemField: 'startLoadingTime',
    label: 'Start Loading Time',
    required: false,
    category: 'Delivery',
    dataType: 'time',
    description: 'Actual time loading operations started (e.g. 08:00 AM)',
    aliases: ['start loading time', 'start loading', 'loading start', 'loading start time', 'started loading'],
    modules: ['dispatch'],
    columnWidth: 20,
    sampleValueDispatch: '08:00 AM',
  },
  {
    excelHeader: 'End Loading Time',
    systemField: 'endLoadingTime',
    label: 'End Loading Time',
    required: false,
    category: 'Delivery',
    dataType: 'time',
    description: 'Actual time loading operations finished (e.g. 09:15 AM)',
    aliases: ['end loading time', 'end loading', 'loading end', 'loading end time', 'loading finished'],
    modules: ['dispatch'],
    columnWidth: 18,
    sampleValueDispatch: '09:15 AM',
  },
  {
    excelHeader: 'Actual Departure',
    systemField: 'actualDepartureTime',
    label: 'Actual Departure Time',
    required: false,
    category: 'Delivery',
    dataType: 'time',
    description: 'Actual time vehicle departed from compound (e.g. 09:45 AM)',
    aliases: ['actual departure', 'actual departure time', 'departure time', 'truck departure', 'time out', 'gate out'],
    modules: ['dispatch'],
    columnWidth: 18,
    sampleValueDispatch: '09:45 AM',
  },
  {
    excelHeader: 'Manifest Number',
    systemField: 'referenceNumber',
    label: 'Manifest / Ref Number',
    required: false,
    category: 'Identification',
    dataType: 'string',
    description: 'Trip manifest or reference identifier',
    aliases: ['manifest number', 'manifest no', 'manifest', 'manifest ref', 'trip number', 'trip no'],
    modules: ['dispatch'],
    columnWidth: 42,
    sampleValueDispatch: 'MNF-2026-0825 [SAMPLE — DELETE BEFORE IMPORT]',
  },
  {
    excelHeader: 'Remarks',
    systemField: 'reasonForDelay',
    label: 'Remarks / Notes',
    required: false,
    category: 'Delivery',
    dataType: 'string',
    description: 'Operational notes or special handling remarks',
    aliases: ['remarks', 'remark', 'notes', 'comments', 'special instructions'],
    modules: ['dispatch'],
    columnWidth: 32,
    sampleValueDispatch: 'SAMPLE — DELETE BEFORE IMPORT',
  },
  {
    excelHeader: 'Planned Delivery Date',
    systemField: 'plannedDeliveryDate',
    label: 'Planned Delivery Date',
    required: false,
    category: 'Delivery',
    dataType: 'date',
    description: 'Target or scheduled delivery arrival date (MM/DD/YYYY or YYYY-MM-DD)',
    aliases: ['planned delivery date', 'planned delivery', 'target delivery date', 'scheduled delivery date', 'scheduled date', 'target date'],
    modules: ['dispatch'],
    columnWidth: 22,
    sampleValueDispatch: '08/28/2026',
  },
];

/**
 * Backward compatibility field definitions list for UI components.
 */
export const OFII_FIELD_DEFINITIONS: OFIIFieldDefinition[] = [
  ...OFII_SHARED_SCHEMA.map((s) => ({
    key: s.systemField,
    label: s.label,
    required: s.required,
    category: s.category,
    description: s.description,
    aliases: [s.excelHeader.toLowerCase(), ...s.aliases],
  })),
  // Additional internal/operational field definitions
  {
    key: 'awbCourierRefNumber',
    label: 'Air Waybill / Courier Ref No.',
    required: false,
    category: 'Identification',
    description: 'Air Waybill, Bill of Lading, or carrier tracking number',
    aliases: ['awb', 'awb number', 'air waybill', 'waybill', 'courier ref', 'courier reference', 'tracking number', 'bl number', 'bill of lading'],
  },
  {
    key: 'unit',
    label: 'Unit',
    required: false,
    category: 'Cargo',
    description: 'Package unit (Boxes, Cases, Pallets, Bundles)',
    aliases: ['unit', 'uom', 'package unit', 'unit of measure'],
  },
  {
    key: 'declaredValue',
    label: 'Declared Value',
    required: false,
    category: 'Cargo',
    description: 'Declared commercial value for insurance',
    aliases: ['declared value', 'cargo value', 'invoice value', 'commercial value'],
  },
  {
    key: 'driverName',
    label: 'Driver Name',
    required: false,
    category: 'Cargo',
    description: 'Assigned truck driver / hauler operator',
    aliases: ['driver', 'driver name', 'assigned driver', 'truck driver', 'operator'],
  },
  {
    key: 'deliveryStatus',
    label: 'Delivery Status',
    required: false,
    category: 'Delivery',
    description: 'Delivered, In Transit, Pending Delivery, or Delayed',
    aliases: ['delivery status', 'status', 'shipment status', 'dispatch status', 'current status'],
  },
  {
    key: 'receiversName',
    label: "Receiver's Name",
    required: false,
    category: 'Delivery',
    description: 'Name of the person who signed at destination',
    aliases: ['receiver', 'receivers name', "receiver's name", 'received by', 'signatory', 'consignee receiver'],
  },
  {
    key: 'deliveryLeadTimeDays',
    label: 'Delivery Leadtime',
    required: false,
    category: 'Delivery',
    description: 'Agreed Delivery SLA in business days (Auto-calculated)',
    aliases: ['delivery leadtime', 'delivery lead time', 'lead time', 'delivery sla'],
  },
  {
    key: 'deliveryTatDays',
    label: 'Delivery TAT',
    required: false,
    category: 'Delivery',
    description: 'Actual Turnaround Time for delivery in days',
    aliases: ['delivery tat', 'delivery turn around time', 'tat days', 'delivery days'],
  },
  {
    key: 'deliveryPerformance',
    label: 'Delivery Performance',
    required: false,
    category: 'Delivery',
    description: 'HIT / MISSED vs delivery lead time SLA',
    aliases: ['delivery performance', 'delivery sla performance', 'sla result', 'delivery hit missed'],
  },
  {
    key: 'podStatus',
    label: 'POD Status',
    required: false,
    category: 'POD',
    description: 'Returned, Pending Return, Transmitted, or Under Review',
    aliases: ['pod status', 'proof of delivery status', 'pod state'],
  },
  {
    key: 'podLeadTimeDays',
    label: 'POD Leadtime',
    required: false,
    category: 'POD',
    description: 'Agreed POD return SLA in business days (default: 3 days)',
    aliases: ['pod leadtime', 'pod lead time', 'pod sla'],
  },
  {
    key: 'podTatDays',
    label: 'POD TAT',
    required: false,
    category: 'POD',
    description: 'Actual POD Return turnaround time in days',
    aliases: ['pod tat', 'pod turn around time', 'pod days'],
  },
  {
    key: 'podPerformance',
    label: 'POD Performance',
    required: false,
    category: 'POD',
    description: 'HIT / MISSED vs POD SLA',
    aliases: ['pod performance', 'pod sla performance', 'pod hit missed'],
  },
  {
    key: 'podReasonForDelay',
    label: 'POD Reason of Delay',
    required: false,
    category: 'POD',
    description: 'Explanation for delayed hardcopy POD submission',
    aliases: ['pod reason of delay', 'pod delay reason', 'pod exception', 'reason for pod delay'],
  },
];

/**
 * Normalizes string for matching (lowercase, strips special chars).
 */
export function cleanHeaderString(str: string): string {
  return str
    .toLowerCase()
    .replace(/[._\-–—/()]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Auto-maps Excel columns to OFII fields using the Shared Column Schema.
 */
export function autoMapHeaders(
  headers: string[],
  targetModule: 'forwarding' | 'dispatch' = 'forwarding'
): Record<string, OFIIFieldKey> {
  const mapping: Record<string, OFIIFieldKey> = {};
  const usedFields = new Set<OFIIFieldKey>();

  headers.forEach((header) => {
    const clean = cleanHeaderString(header);
    if (!clean) {
      mapping[header] = 'none';
      return;
    }

    let matchedField: OFIIFieldKey = 'none';

    // 1. Direct match with exact Excel Header from Shared Schema
    for (const schemaCol of OFII_SHARED_SCHEMA) {
      if (usedFields.has(schemaCol.systemField)) continue;
      if (clean === cleanHeaderString(schemaCol.excelHeader)) {
        matchedField = schemaCol.systemField;
        break;
      }
    }

    // 2. Target-module specific priority checks if not yet matched
    if (matchedField === 'none') {
      if (targetModule === 'dispatch') {
        if (clean === 'planned delivery date' || clean === 'planned delivery' || clean === 'target delivery date' || clean === 'target delivery' || clean === 'scheduled date') {
          matchedField = 'plannedDeliveryDate';
        } else if (clean === 'destination' || clean === 'destination address' || clean === 'delivery destination' || clean === 'place') {
          matchedField = 'destination';
        } else if (clean === 'plate number' || clean === 'truck plate' || clean === 'plate no' || clean === 'plate no.') {
          matchedField = 'plateNumber';
        } else if (clean === 'driver' || clean === 'driver name' || clean === 'assigned driver') {
          matchedField = 'driverName';
        } else if (clean === 'delivery date' || clean === 'dispatch date' || clean === 'date dispatched' || clean === 'date') {
          matchedField = 'actualDispatchDate';
        } else if (clean === 'time arrived' || clean === 'arrival time' || clean === 'truck arrival' || clean === 'truck arrival time' || clean === 'time in' || clean === 'gate in') {
          matchedField = 'timeArrived';
        } else if (clean === 'start loading' || clean === 'start loading time' || clean === 'loading start' || clean === 'loading start time' || clean === 'start load') {
          matchedField = 'startLoadingTime';
        } else if (clean === 'end loading' || clean === 'end loading time' || clean === 'loading end' || clean === 'loading end time' || clean === 'loading finished' || clean === 'end load') {
          matchedField = 'endLoadingTime';
        } else if (clean === 'actual departure' || clean === 'actual departure time' || clean === 'departure time' || clean === 'truck departure' || clean === 'departed time' || clean === 'time out' || clean === 'gate out') {
          matchedField = 'actualDepartureTime';
        }
      }
    }

    // 3. Exact match with label or aliases in OFII_SHARED_SCHEMA & OFII_FIELD_DEFINITIONS
    if (matchedField === 'none') {
      for (const def of OFII_FIELD_DEFINITIONS) {
        if (usedFields.has(def.key)) continue;

        if (clean === cleanHeaderString(def.label)) {
          matchedField = def.key;
          break;
        }

        for (const alias of def.aliases) {
          if (clean === cleanHeaderString(alias)) {
            matchedField = def.key;
            break;
          }
        }

        if (matchedField !== 'none') break;
      }
    }

    // 4. Fuzzy / substring matches if still not matched
    if (matchedField === 'none') {
      for (const def of OFII_FIELD_DEFINITIONS) {
        if (usedFields.has(def.key)) continue;

        for (const alias of def.aliases) {
          const cleanAlias = cleanHeaderString(alias);
          if (clean.includes(cleanAlias) || cleanAlias.includes(clean)) {
            if (clean.includes('dispatch') && def.key === 'actualDeliveryDate') continue;
            if (clean.includes('planned') && def.key === 'actualDeliveryDate') continue;
            if (clean.includes('planned') && def.key === 'actualDispatchDate') continue;
            if (clean.includes('pod') && clean.includes('return') && def.key === 'podNumber') continue;

            matchedField = def.key;
            break;
          }
        }
        if (matchedField !== 'none') break;
      }
    }

    if (matchedField !== 'none') {
      usedFields.add(matchedField);
      mapping[header] = matchedField;
    } else {
      mapping[header] = 'none';
    }
  });

  return mapping;
}

/**
 * Parses dates from Excel (handles numeric serials, ISO strings, slash/dash dates).
 * Safely returns YYYY-MM-DD string or empty string if invalid/empty.
 */
export function parseExcelDate(val: any): string {
  if (val === null || val === undefined || val === '') return '';

  if (val instanceof Date) {
    if (!isNaN(val.getTime())) {
      return val.toISOString().split('T')[0];
    }
    return '';
  }

  // Handle Excel Serial Number (e.g. 45894)
  if (typeof val === 'number') {
    if (isNaN(val) || val <= 0) return '';
    // 25569 is days between 1900-01-01 and 1970-01-01
    const date = new Date((val - 25569) * 86400 * 1000);
    if (!isNaN(date.getTime())) {
      return date.toISOString().split('T')[0];
    }
    return '';
  }

  const str = String(val).trim();
  if (!str) return '';

  // Check YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    const testDate = new Date(str);
    if (!isNaN(testDate.getTime())) {
      return str;
    }
    return '';
  }

  // Handle MM/DD/YYYY, DD/MM/YYYY, or YYYY/MM/DD
  const parts = str.split(/[/.-]/);
  if (parts.length === 3) {
    let year = parts[2];
    let month = parts[0];
    let day = parts[1];

    if (parts[0].length === 4) {
      year = parts[0];
      month = parts[1];
      day = parts[2];
    } else if (year.length === 2) {
      year = `20${year}`;
    }

    const numM = parseInt(month, 10);
    const numD = parseInt(day, 10);
    const numY = parseInt(year, 10);

    if (numY > 1900 && numM >= 1 && numM <= 12 && numD >= 1 && numD <= 31) {
      const m = String(numM).padStart(2, '0');
      const d = String(numD).padStart(2, '0');
      const iso = `${numY}-${m}-${d}`;
      const testDate = new Date(iso);
      if (!isNaN(testDate.getTime())) {
        return iso;
      }
    }
  }

  // Standard Date parse fallback
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    try {
      const y = parsed.getFullYear();
      const m = String(parsed.getMonth() + 1).padStart(2, '0');
      const d = String(parsed.getDate()).padStart(2, '0');
      if (y > 1900 && y < 2200) {
        return `${y}-${m}-${d}`;
      }
    } catch (e) {
      return '';
    }
  }

  return '';
}

/**
 * Cleans numbers from Excel cells (currency symbols, commas, whitespace).
 */
export function parseExcelNumber(val: any, defaultVal = 0): number {
  if (typeof val === 'number') return isNaN(val) ? defaultVal : val;
  if (!val) return defaultVal;
  const str = String(val).replace(/[^0-9.-]/g, '');
  const num = parseFloat(str);
  return isNaN(num) ? defaultVal : num;
}

/**
 * Reads an uploaded Excel file (.xlsx or .xls) and extracts raw headers & rows.
 * Automatically chooses the appropriate worksheet and filters out sample template rows.
 */
export async function readExcelFile(
  file: File,
  targetModule: 'forwarding' | 'dispatch' = 'forwarding'
): Promise<{
  headers: string[];
  rows: Record<string, any>[];
  fileName: string;
  fileSizeFormatted: string;
  sampleRowsExcludedCount?: number;
}> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, {
          type: 'array',
          cellDates: true,
          cellNF: false,
          cellText: false,
        });

        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          throw new Error('No worksheets found in this Excel file.');
        }

        // Intelligently select worksheet
        let selectedSheetName = workbook.SheetNames[0];

        // Filter out instruction sheets unless it's the only one
        const dataSheets = workbook.SheetNames.filter(
          (name) => !/instruction|guide|readme|reference|notes/i.test(name.trim())
        );
        const candidates = dataSheets.length > 0 ? dataSheets : workbook.SheetNames;

        if (targetModule === 'dispatch') {
          const dispatchMatch = candidates.find((name) =>
            /dispatch|daily|delivery/i.test(name)
          );
          if (dispatchMatch) selectedSheetName = dispatchMatch;
          else selectedSheetName = candidates[0];
        } else {
          const forwardingMatch = candidates.find((name) =>
            /forwarding|progressive|shipment|report/i.test(name)
          );
          if (forwardingMatch) selectedSheetName = forwardingMatch;
          else selectedSheetName = candidates[0];
        }

        const worksheet = workbook.Sheets[selectedSheetName];
        if (!worksheet) {
          throw new Error(`Unable to read worksheet "${selectedSheetName}".`);
        }

        // Convert to JSON array of objects
        const rawJson: any[] = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: '',
          blankrows: false,
        });

        if (!rawJson || rawJson.length === 0) {
          throw new Error(`Worksheet "${selectedSheetName}" is empty. No records found.`);
        }

        // Find header row (first non-empty row)
        let headerRowIndex = -1;
        for (let i = 0; i < rawJson.length; i++) {
          const row = rawJson[i];
          if (Array.isArray(row) && row.some((cell) => cell !== undefined && String(cell).trim() !== '')) {
            headerRowIndex = i;
            break;
          }
        }

        if (headerRowIndex === -1) {
          throw new Error('No valid header row was found in this worksheet.');
        }

        const rawHeaderRow = rawJson[headerRowIndex] as any[];
        const headers: string[] = rawHeaderRow
          .map((h, idx) => (h !== undefined && String(h).trim() !== '' ? String(h).trim() : `Column_${idx + 1}`))
          .filter(Boolean);

        const rows: Record<string, any>[] = [];
        let sampleRowsExcludedCount = 0;

        for (let r = headerRowIndex + 1; r < rawJson.length; r++) {
          const rowArray = rawJson[r] as any[];
          if (!rowArray || !Array.isArray(rowArray)) continue;

          // Skip completely blank rows
          const hasData = rowArray.some((cell) => cell !== undefined && String(cell).trim() !== '');
          if (!hasData) continue;

          const rowObj: Record<string, any> = {};
          headers.forEach((header, colIdx) => {
            rowObj[header] = rowArray[colIdx] !== undefined ? rowArray[colIdx] : '';
          });

          // Check if row is a sample row (contains "SAMPLE" or "DELETE BEFORE IMPORT" in any cell)
          const isSampleRow = Object.values(rowObj).some((val) => {
            if (typeof val !== 'string') return false;
            const upper = val.toUpperCase();
            return (
              upper.includes('SAMPLE — DELETE BEFORE IMPORT') ||
              upper.includes('SAMPLE - DELETE BEFORE IMPORT') ||
              upper.includes('DELETE BEFORE IMPORT') ||
              upper.includes('[SAMPLE') ||
              upper.includes('SAMPLE ROW') ||
              (upper.includes('SAMPLE') && (upper.includes('PRJ') || upper.includes('MNF') || upper.includes('POD')))
            );
          });

          if (isSampleRow) {
            sampleRowsExcludedCount++;
            continue; // Exclude sample row from actual import records
          }

          rows.push(rowObj);
        }

        if (rows.length === 0 && sampleRowsExcludedCount > 0) {
          throw new Error(
            'The uploaded Excel file only contains the sample template row. Please add your actual shipment records below the header row and re-upload.'
          );
        }

        if (rows.length === 0) {
          throw new Error('No shipment records were found in this file.');
        }

        // Format file size
        const sizeKb = (file.size / 1024).toFixed(1);
        const fileSizeFormatted = file.size > 1024 * 1024 
          ? `${(file.size / (1024 * 1024)).toFixed(2)} MB` 
          : `${sizeKb} KB`;

        resolve({
          headers,
          rows,
          fileName: file.name,
          fileSizeFormatted,
          sampleRowsExcludedCount,
        });
      } catch (err: any) {
        reject(new Error(err.message || 'Unable to read this Excel file.'));
      }
    };

    reader.onerror = () => {
      reject(new Error('Unable to read this Excel file.'));
    };

    reader.readAsArrayBuffer(file);
  });
}

/**
 * Generates and triggers download of the standardized OFII Monitoring Bulk Import Excel Template.
 * Built directly from the Unified OFII_SHARED_SCHEMA to guarantee 100% template-to-importer compatibility.
 */
export function downloadOFIIExcelTemplate(
  preferredSheet?: 'forwarding' | 'dispatch' | 'all'
): void {
  const workbook = XLSX.utils.book_new();

  // =========================================================================
  // 1. WORKSHEET: FORWARDING PROGRESSIVE REPORT
  // =========================================================================
  const forwardingCols = OFII_SHARED_SCHEMA.filter((col) => col.modules.includes('forwarding'));
  const forwardingHeaders = forwardingCols.map((col) => col.excelHeader);
  const forwardingSampleRow: Record<string, any> = {};
  forwardingCols.forEach((col) => {
    forwardingSampleRow[col.excelHeader] = col.sampleValueForwarding ?? '';
  });

  const wsForwarding = XLSX.utils.json_to_sheet([forwardingSampleRow], { header: forwardingHeaders });
  wsForwarding['!cols'] = forwardingCols.map((col) => ({ wch: col.columnWidth }));

  // Valid clients for dropdowns & validation
  const validClientsList = [
    'Philippine Charity Sweepstakes Office',
    'Golden Archers Development Corporation',
    'Alexandria and Centers of Wisdom Corporation',
    'Oriental Merchants',
    'Vamsler Philippines',
    'RefaMED',
    'Dunsk Kuhner',
    'Intelligent Skin Care Inc.',
  ];
  const validModesList = ['Air', 'Land', 'Sea', 'RORO'];
  const validAreasList = ['Luzon', 'Visayas', 'Mindanao'];

  // Add Excel Data Validations (Dropdown lists) to Forwarding Sheet
  // Column C (index 2) = Client Name, Column D (index 3) = Mode of Shipment, Column E (index 4) = Delivery Area
  const forwardingValidations = [
    {
      sqref: 'C2:C500',
      type: 'list',
      allowBlank: false,
      showErrorMessage: true,
      errorTitle: 'Invalid Client',
      error: 'Please select a valid registered client from the dropdown list.',
      formulae: [`"${validClientsList.join(',')}"`],
    },
    {
      sqref: 'D2:D500',
      type: 'list',
      allowBlank: false,
      showErrorMessage: true,
      errorTitle: 'Invalid Mode of Shipment',
      error: 'Please select an allowed Mode of Shipment: Air, Land, Sea, or RORO.',
      formulae: [`"${validModesList.join(',')}"`],
    },
    {
      sqref: 'E2:E500',
      type: 'list',
      allowBlank: false,
      showErrorMessage: true,
      errorTitle: 'Invalid Delivery Area',
      error: 'Please select an allowed Delivery Area: Luzon, Visayas, or Mindanao.',
      formulae: [`"${validAreasList.join(',')}"`],
    },
  ];
  (wsForwarding as any)['!dataValidation'] = forwardingValidations;
  (wsForwarding as any)['!dataValidations'] = forwardingValidations;

  // =========================================================================
  // 2. WORKSHEET: DAILY DISPATCH MONITORING
  // =========================================================================
  const dispatchCols = OFII_SHARED_SCHEMA.filter((col) => col.modules.includes('dispatch'));
  const dispatchHeaders = dispatchCols.map((col) => col.excelHeader);
  const dispatchSampleRow: Record<string, any> = {};
  dispatchCols.forEach((col) => {
    dispatchSampleRow[col.excelHeader] = col.sampleValueDispatch ?? '';
  });

  const wsDispatch = XLSX.utils.json_to_sheet([dispatchSampleRow], { header: dispatchHeaders });
  wsDispatch['!cols'] = dispatchCols.map((col) => ({ wch: col.columnWidth }));

  // =========================================================================
  // 3. WORKSHEET: INSTRUCTIONS & REFERENCE GUIDE
  // =========================================================================
  const instructionsData = [
    ['ORIENT FREIGHT INTERNATIONAL, INC. (OFII) - BULK IMPORT EXCEL TEMPLATE'],
    ['Standardized Excel Template for Offline Bulk Encoding and Automated SLA Calculations'],
    [],
    ['================================================================================================='],
    ['INSTRUCTIONS & ENCODING GUIDELINES (10 STEPS):'],
    ['================================================================================================='],
    ['Step 1', 'Fill in the required columns for each shipment or dispatch record.'],
    ['Step 2', 'Required Forwarding columns: Client Name, Delivery Area, Mode of Shipment, Reference Number, Consignee, Actual Dispatched Date.'],
    ['Step 3', 'Do NOT rename, remove, or modify any column header titles in Row 1.'],
    ['Step 4', 'Select Client Name from the Excel dropdown or use exact registered names from the table below.'],
    ['Step 5', 'Select Delivery Area: "Luzon", "Visayas", or "Mindanao".'],
    ['Step 6', 'Select Mode of Shipment: "Air", "Land", "Sea", or "RORO".'],
    ['Step 7', 'Coordinator is auto-assigned from Client Name (e.g., Intelligent Skin Care Inc. -> Rojaylene Baldesancho).'],
    ['Step 8', 'Delivery Leadtime is auto-calculated from Client + Mode + Delivery Area using official SLA working days.'],
    ['Step 9', 'Delete or replace the marked [SAMPLE — DELETE BEFORE IMPORT] row before importing.'],
    ['Step 10', 'Save your completed file as .xlsx, return to OFII Monitoring System, and click [ IMPORT EXCEL ].'],
    [],
    ['================================================================================================='],
    ['REGISTERED CLIENTS & ASSIGNED ACCOUNT COORDINATORS REFERENCE:'],
    ['================================================================================================='],
    ['Client Name', 'Client Code', 'Assigned Account Coordinator', 'Notes'],
    ['Philippine Charity Sweepstakes Office', 'PCSO-001', 'Alodia Manalansan', 'Auto-assigned by system if blank'],
    ['Golden Archers Development Corporation', 'GADC-002', 'Justine Ryan Paular', 'Auto-assigned by system if blank'],
    ['Alexandria and Centers of Wisdom Corporation', 'ACWC-003', 'Alodia Manalansan', 'Auto-assigned by system if blank'],
    ['Oriental Merchants', 'OM-004', 'Alodia Manalansan', 'Auto-assigned by system if blank'],
    ['Vamsler Philippines', 'VAM-005', 'Alodia Manalansan', 'Auto-assigned by system if blank'],
    ['RefaMED', 'RFM-006', 'Alodia Manalansan', 'Auto-assigned by system if blank'],
    ['Dunsk Kuhner', 'DK-007', 'Alodia Manalansan', 'Auto-assigned by system if blank'],
    ['Intelligent Skin Care Inc.', 'ISCI-008', 'Rojaylene Baldesancho', 'Auto-assigned by system if blank'],
    [],
    ['================================================================================================='],
    ['ACCEPTED VALUES REFERENCE:'],
    ['================================================================================================='],
    ['Mode of Shipment', 'Air, Land, Sea, RORO (or Air Freight, Land Freight, Sea Freight)'],
    ['Delivery Area', 'Luzon, Visayas, Mindanao'],
    ['Date Format', 'MM/DD/YYYY (e.g., 08/25/2026) or YYYY-MM-DD (e.g., 2026-08-25)'],
    ['Time Format', 'HH:MM AM/PM (e.g., 07:30 AM, 02:45 PM) or 24-Hour (e.g., 07:30, 14:45)'],
    [],
    ['================================================================================================='],
    ['AUTOMATED SYSTEM CALCULATIONS (DO NOT ENCODE FORMULAS IN EXCEL):'],
    ['================================================================================================='],
    ['The OFII Leadtime & SLA Engine automatically computes upon import:'],
    ['1. Delivery Leadtime (Working days SLA excluding Sundays and Philippine regular/special non-working holidays)'],
    ['   Example: Intelligent Skin Care Inc. + Air + Luzon = 5 working days SLA'],
    ['2. Expected Delivery Date (Dispatched Date + SLA business days)'],
    ['3. Delivery Performance (HIT vs MISSED based on actual delivery vs SLA / RDD)'],
    ['4. POD Return Due Date (Actual Delivery Date + 3 business days SLA)'],
    ['5. POD SLA & Status (Returned, Pending Return, SLA Breached alerts)'],
    ['6. Automatic Client Name -> Coordinator mapping (e.g., Intelligent Skin Care Inc. -> Rojaylene Baldesancho)'],
  ];

  const wsInstructions = XLSX.utils.aoa_to_sheet(instructionsData);
  wsInstructions['!cols'] = [
    { wch: 42 },
    { wch: 65 },
    { wch: 32 },
    { wch: 35 },
  ];

  // Append sheets in order
  if (preferredSheet === 'dispatch') {
    XLSX.utils.book_append_sheet(workbook, wsDispatch, 'Daily Dispatch Monitoring');
    XLSX.utils.book_append_sheet(workbook, wsForwarding, 'Forwarding Progressive Report');
  } else {
    XLSX.utils.book_append_sheet(workbook, wsForwarding, 'Forwarding Progressive Report');
    XLSX.utils.book_append_sheet(workbook, wsDispatch, 'Daily Dispatch Monitoring');
  }
  XLSX.utils.book_append_sheet(workbook, wsInstructions, 'Instructions');

  // Trigger download
  XLSX.writeFile(workbook, 'OFII_Monitoring_Import_Template.xlsx');
}

/**
 * Backward compatibility alias for downloadSampleExcelTemplate
 */
export function downloadSampleExcelTemplate(
  targetModule: 'forwarding' | 'dispatch' = 'forwarding'
): void {
  downloadOFIIExcelTemplate(targetModule);
}
