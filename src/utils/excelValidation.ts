import { 
  ForwardingProgressiveRecord, 
  ClientSummary, 
  DispatchRecord,
  ForwardingMode, 
  PhilippineArea, 
  ForwardingDeliveryStatus, 
  PODStatus,
  PerformanceResult,
  OFIIFieldKey 
} from '../types';
import { 
  getAutoDeliveryLeadTime, 
  computeDeliveryPerformance, 
  computePodPerformance,
  calculatePodReturnDueDate,
  determineAutomaticPodStatus,
} from './forwardingCalculations';
import { calculateExpectedDeliveryDate } from './leadtimeEngine';
import { 
  getClientAssignedCoordinator, 
  normalizeClientString, 
  STANDARD_CLIENT_ALIASES 
} from './dataSync';
import { parseExcelDate, parseExcelNumber } from './excelParser';

export type RowValidationStatus = 'VALID' | 'WARNING' | 'INVALID' | 'DUPLICATE';

export type CellErrorType = 
  | 'MISSING_REQUIRED' 
  | 'INVALID_VALUE' 
  | 'INVALID_DATE' 
  | 'INVALID_TIME' 
  | 'INVALID_NUMBER';

export interface CellValidationError {
  fieldKey: string;
  fieldName: string;
  errorType: CellErrorType;
  message: string;
  currentValue?: any;
}

export interface ValidatedImportRow {
  rowIndex: number;
  originalRow: Record<string, any>;
  mappedRecord: ForwardingProgressiveRecord;
  status: RowValidationStatus;
  errors: string[];
  warnings: string[];
  cellErrors: Record<string, CellValidationError>;
  cellWarnings: Record<string, string>;
  isDuplicate: boolean;
  duplicateReason?: string;
  isClientRecognized: boolean;
  recognizedClient?: ClientSummary;
  assignedCoordinator: string;
  isSelectedForImport: boolean;
  duplicateAction: 'skip' | 'import_anyway';
}

export interface ValidationSummary {
  totalRows: number;
  validCount: number;
  warningCount: number;
  invalidCount: number;
  duplicateCount: number;
  readyToImportCount: number;
  rows: ValidatedImportRow[];
}

/**
 * Normalizes Mode of Shipment from loose Excel text.
 */
export function normalizeModeOfShipment(raw: any): { mode: ForwardingMode; isRecognized: boolean } {
  if (!raw) return { mode: 'Land Freight', isRecognized: false };
  const str = String(raw).trim().toLowerCase();

  if (str.includes('roro') || str.includes('roll-on') || str.includes('roll on')) {
    return { mode: 'RORO', isRecognized: true };
  }
  if (str.includes('air') || str.includes('plane') || str.includes('flight')) {
    return { mode: 'Air Freight', isRecognized: true };
  }
  if (str.includes('sea') || str.includes('vessel') || str.includes('boat') || str.includes('marine') || str.includes('ocean')) {
    return { mode: 'Sea Freight', isRecognized: true };
  }
  if (str.includes('land') || str.includes('truck') || str.includes('road') || str.includes('ground') || str.includes('ftl') || str.includes('ltl')) {
    return { mode: 'Land Freight', isRecognized: true };
  }

  return { mode: 'Land Freight', isRecognized: false };
}

/**
 * Normalizes Area from loose Excel text.
 */
export function normalizeArea(raw: any): { area: PhilippineArea; isRecognized: boolean } {
  if (!raw) return { area: 'Luzon', isRecognized: false };
  const str = String(raw).trim().toLowerCase();

  if (str.includes('ncr') || str.includes('metro manila') || str.includes('manila')) {
    return { area: 'NCR', isRecognized: true };
  }
  if (str.includes('vis') || str.includes('cebu') || str.includes('iloilo') || str.includes('bacolod') || str.includes('tacloban') || str.includes('bohol')) {
    return { area: 'Visayas', isRecognized: true };
  }
  if (str.includes('min') || str.includes('davao') || str.includes('cagayan') || str.includes('cdo') || str.includes('gensan') || str.includes('zamboanga')) {
    return { area: 'Mindanao', isRecognized: true };
  }
  if (str.includes('luz') || str.includes('pampanga') || str.includes('batangas') || str.includes('cavite') || str.includes('laguna') || str.includes('baguio') || str.includes('north') || str.includes('south')) {
    return { area: 'Luzon', isRecognized: true };
  }

  return { area: 'Luzon', isRecognized: false };
}

/**
 * Normalizes Delivery Status.
 */
export function normalizeDeliveryStatus(raw: any, hasActualDeliveryDate: boolean): ForwardingDeliveryStatus {
  if (!raw && hasActualDeliveryDate) return 'Delivered';
  if (!raw) return 'In Transit';

  const str = String(raw).trim().toLowerCase();
  if (str.includes('deliv') || str.includes('received') || str.includes('complete') || str.includes('done')) {
    return 'Delivered';
  }
  if (str.includes('delay') || str.includes('late') || str.includes('slipped') || str.includes('failed')) {
    return 'Delayed';
  }
  if (str.includes('transit') || str.includes('route') || str.includes('shipping') || str.includes('dispatched')) {
    return 'In Transit';
  }
  if (str.includes('pending') || str.includes('booked') || str.includes('queue')) {
    return 'Pending Delivery';
  }

  return hasActualDeliveryDate ? 'Delivered' : 'In Transit';
}

/**
 * Finds a matching client from the shared client dataset using robust normalization and alias dictionaries.
 */
export function matchClient(
  rawClientName: string,
  clients: ClientSummary[]
): { client: ClientSummary | null; isRecognized: boolean } {
  if (!rawClientName) return { client: null, isRecognized: false };
  const clean = normalizeClientString(rawClientName);
  if (!clean) return { client: null, isRecognized: false };

  // 1. Direct name, ID, or code match
  const directMatch = clients.find(
    (c) =>
      c.id === rawClientName ||
      normalizeClientString(c.name) === clean ||
      normalizeClientString(c.code) === clean
  );
  if (directMatch) return { client: directMatch, isRecognized: true };

  // 2. Standard Alias match
  const canonicalName = STANDARD_CLIENT_ALIASES[clean];
  if (canonicalName) {
    const aliasMatch = clients.find(
      (c) => normalizeClientString(c.name) === normalizeClientString(canonicalName)
    );
    if (aliasMatch) return { client: aliasMatch, isRecognized: true };
  }

  // 3. Substring / partial match
  const partialMatch = clients.find((c) => {
    const cClean = normalizeClientString(c.name);
    return cClean.includes(clean) || clean.includes(cClean);
  });
  if (partialMatch) return { client: partialMatch, isRecognized: true };

  return { client: null, isRecognized: false };
}

/**
 * Validates a single import record with cell-level error tracking, auto-fill capabilities, and SLA calculations.
 */
export function validateSingleRow({
  rowIndex,
  originalRow,
  mappedRecord,
  clients,
  existingRefs,
  existingPods,
  existingAwbs,
  batchRefs,
  batchPods,
  targetModule = 'forwarding',
  duplicateAction = 'skip',
}: {
  rowIndex: number;
  originalRow: Record<string, any>;
  mappedRecord: Partial<ForwardingProgressiveRecord>;
  clients: ClientSummary[];
  existingRefs?: Set<string>;
  existingPods?: Set<string>;
  existingAwbs?: Set<string>;
  batchRefs?: Map<string, number>;
  batchPods?: Map<string, number>;
  targetModule?: 'forwarding' | 'dispatch';
  duplicateAction?: 'skip' | 'import_anyway';
}): ValidatedImportRow {
  const errors: string[] = [];
  const warnings: string[] = [];
  const cellErrors: Record<string, CellValidationError> = {};
  const cellWarnings: Record<string, string> = {};

  const rawClient = String(mappedRecord.client || '').trim();
  let clientName = rawClient;
  let clientId: string | undefined = mappedRecord.clientId;
  let coordinator = mappedRecord.coordinator || 'Alodia Manalansan';
  let isClientRecognized = false;
  let recognizedClient: ClientSummary | undefined = undefined;

  // 1. Client Validation
  if (!rawClient) {
    const msg = 'Client Name is required';
    errors.push(msg);
    cellErrors['client'] = {
      fieldKey: 'client',
      fieldName: 'Client Name',
      errorType: 'MISSING_REQUIRED',
      message: msg,
      currentValue: '',
    };
  } else {
    const matchRes = matchClient(rawClient, clients);
    if (matchRes.isRecognized && matchRes.client) {
      isClientRecognized = true;
      recognizedClient = matchRes.client;
      clientName = matchRes.client.name;
      clientId = matchRes.client.id;
      coordinator = matchRes.client.assignedCoordinator || matchRes.client.accountManager || 'Alodia Manalansan';
    } else {
      isClientRecognized = false;
      const warnMsg = `Unrecognized Client: "${rawClient}". Please select from registered clients.`;
      warnings.push(warnMsg);
      cellWarnings['client'] = warnMsg;
      coordinator = getClientAssignedCoordinator(clients, rawClient, 'Alodia Manalansan');
    }
  }

  // 2. Consignee Validation
  const consignee = String(mappedRecord.consignee || '').trim();
  if (!consignee) {
    const msg = 'Consignee / Recipient is required';
    errors.push(msg);
    cellErrors['consignee'] = {
      fieldKey: 'consignee',
      fieldName: 'Consignee',
      errorType: 'MISSING_REQUIRED',
      message: msg,
      currentValue: '',
    };
  }

  // 3. Mode of Shipment Validation
  const rawMode = mappedRecord.modeOfShipment;
  const modeRes = normalizeModeOfShipment(rawMode);
  const modeOfShipment: ForwardingMode = modeRes.mode;
  if (!rawMode) {
    if (targetModule === 'dispatch') {
      warnings.push('Mode of Shipment not specified; defaulted to "Land Freight".');
    } else {
      const msg = 'Mode of Shipment is required (Air, Land, Sea, or RORO)';
      errors.push(msg);
      cellErrors['modeOfShipment'] = {
        fieldKey: 'modeOfShipment',
        fieldName: 'Mode of Shipment',
        errorType: 'MISSING_REQUIRED',
        message: msg,
        currentValue: '',
      };
    }
  } else if (!modeRes.isRecognized) {
    const msg = 'Must be Air, Land, Sea, or RORO';
    cellWarnings['modeOfShipment'] = msg;
    warnings.push(`Unrecognized Mode of Shipment "${rawMode}", defaulted to "${modeOfShipment}"`);
  }

  // 4. Area Validation
  const rawArea = mappedRecord.area;
  const areaRes = normalizeArea(rawArea);
  const area: PhilippineArea = areaRes.area;
  if (!rawArea) {
    if (targetModule === 'dispatch') {
      warnings.push('Delivery Area not specified; defaulted to "Luzon".');
    } else {
      const msg = 'Delivery Area is required (Luzon, Visayas, or Mindanao)';
      errors.push(msg);
      cellErrors['area'] = {
        fieldKey: 'area',
        fieldName: 'Delivery Area',
        errorType: 'MISSING_REQUIRED',
        message: msg,
        currentValue: '',
      };
    }
  } else if (!areaRes.isRecognized) {
    const msg = 'Must be Luzon, Visayas, or Mindanao';
    cellWarnings['area'] = msg;
    warnings.push(`Unrecognized Area "${rawArea}", defaulted to "${area}"`);
  }

  // 5. Reference Number Validation
  const rawRef = String(mappedRecord.referenceNumber || '').trim();
  let referenceNumber = rawRef;
  if (!rawRef) {
    if (targetModule === 'dispatch') {
      const clientPrefix = clientName ? clientName.replace(/[^a-zA-Z]/g, '').slice(0, 4).toUpperCase() : 'OFII';
      referenceNumber = `PRJ-${clientPrefix}-${Math.floor(100 + Math.random() * 900)}`;
      warnings.push(`Auto-generated reference number: ${referenceNumber}`);
    } else {
      const msg = 'Reference Number is required';
      errors.push(msg);
      cellErrors['referenceNumber'] = {
        fieldKey: 'referenceNumber',
        fieldName: 'Reference Number',
        errorType: 'MISSING_REQUIRED',
        message: msg,
        currentValue: '',
      };
    }
  }

  // 6. Actual Dispatched Date Validation
  const rawDispatchDate = mappedRecord.actualDispatchDate;
  let actualDispatchDate = parseExcelDate(rawDispatchDate);
  const plannedDeliveryDate = parseExcelDate(mappedRecord.plannedDeliveryDate);

  if (!rawDispatchDate) {
    if (plannedDeliveryDate) {
      actualDispatchDate = plannedDeliveryDate;
      warnings.push('Dispatch date set from Planned Delivery Date.');
    } else if (targetModule === 'dispatch') {
      actualDispatchDate = '2026-08-25';
      warnings.push('Dispatch date defaulted to current operational date (2026-08-25).');
    } else {
      const msg = 'Dispatch Date is required';
      errors.push(msg);
      cellErrors['actualDispatchDate'] = {
        fieldKey: 'actualDispatchDate',
        fieldName: 'Dispatch Date',
        errorType: 'MISSING_REQUIRED',
        message: msg,
        currentValue: '',
      };
    }
  } else if (!actualDispatchDate) {
    const msg = 'Invalid date format (use MM/DD/YYYY or YYYY-MM-DD)';
    errors.push(msg);
    cellErrors['actualDispatchDate'] = {
      fieldKey: 'actualDispatchDate',
      fieldName: 'Dispatch Date',
      errorType: 'INVALID_DATE',
      message: msg,
      currentValue: rawDispatchDate,
    };
  }

  // Quantity check
  let quantity = 100;
  if (mappedRecord.quantity !== undefined && (mappedRecord.quantity as any) !== '') {
    const parsedQ = parseExcelNumber(mappedRecord.quantity, -1);
    if (parsedQ < 0 || isNaN(parsedQ)) {
      const msg = 'Quantity must be a valid positive number';
      errors.push(msg);
      cellErrors['quantity'] = {
        fieldKey: 'quantity',
        fieldName: 'Quantity',
        errorType: 'INVALID_NUMBER',
        message: msg,
        currentValue: mappedRecord.quantity,
      };
    } else {
      quantity = parsedQ;
    }
  }

  // Parse other dates
  const requestDeliveryDate = parseExcelDate(mappedRecord.requestDeliveryDate);
  const actualDeliveryDate = parseExcelDate(mappedRecord.actualDeliveryDate);
  const dateOfPodReturn = parseExcelDate(mappedRecord.dateOfPodReturn);
  const cbm = mappedRecord.cbm !== undefined && (mappedRecord.cbm as any) !== '' ? parseExcelNumber(mappedRecord.cbm, 0) : undefined;
  const actualWeightKg = mappedRecord.actualWeightKg !== undefined && (mappedRecord.actualWeightKg as any) !== ''
    ? parseExcelNumber(mappedRecord.actualWeightKg, quantity * 12)
    : quantity * 12;
  const volumeWeightKg = mappedRecord.volumeWeightKg !== undefined && (mappedRecord.volumeWeightKg as any) !== ''
    ? parseExcelNumber(mappedRecord.volumeWeightKg, 0)
    : undefined;

  // SLA and Performance Calculations
  const deliveryLeadTimeDays = getAutoDeliveryLeadTime(clientName, modeOfShipment, area);
  let deliveryTatDays = 0;
  let deliveryPerformance: PerformanceResult = 'PENDING';
  if (actualDispatchDate && actualDeliveryDate) {
    const perfRes = computeDeliveryPerformance(
      actualDispatchDate,
      actualDeliveryDate,
      deliveryLeadTimeDays,
      undefined,
      requestDeliveryDate
    );
    deliveryTatDays = perfRes.tatDays;
    deliveryPerformance = perfRes.performance;
  }

  const podDueRes = calculatePodReturnDueDate(actualDeliveryDate, clientName, area);
  const autoPodRes = determineAutomaticPodStatus({
    actualDeliveryDate,
    podReturnDueDate: podDueRes.podReturnDueDate,
    actualPodReturnDate: dateOfPodReturn,
    clientName,
    deliveryArea: area,
  });
  const podTatDays = autoPodRes.podTatDays;
  const podPerformance = autoPodRes.podPerformance;

  // Duplicate checks
  let isDuplicate = false;
  let duplicateReason: string | undefined = undefined;

  if (rawRef && existingRefs?.has(String(rawRef).trim().toLowerCase())) {
    isDuplicate = true;
    duplicateReason = `Reference Number "${rawRef}" already exists in the system.`;
  } else if (mappedRecord.podNumber && existingPods?.has(String(mappedRecord.podNumber).trim().toLowerCase())) {
    isDuplicate = true;
    duplicateReason = `POD Number "${mappedRecord.podNumber}" already exists in the system.`;
  } else if (mappedRecord.awbCourierRefNumber && existingAwbs?.has(String(mappedRecord.awbCourierRefNumber).trim().toLowerCase())) {
    isDuplicate = true;
    duplicateReason = `AWB / Courier Ref "${mappedRecord.awbCourierRefNumber}" already exists in the system.`;
  }

  if (rawRef && batchRefs) {
    const refKey = String(rawRef).trim().toLowerCase();
    const prevIndex = batchRefs.get(refKey);
    if (prevIndex !== undefined && prevIndex !== rowIndex - 1) {
      isDuplicate = true;
      duplicateReason = `Duplicate Reference Number "${rawRef}" detected in Row ${prevIndex + 1} of this Excel file.`;
    }
  }

  const expectedLeadtimeRes = calculateExpectedDeliveryDate(actualDispatchDate, modeOfShipment, area);

  // Determine Row Status
  let status: RowValidationStatus = 'VALID';
  if (Object.keys(cellErrors).length > 0 || errors.length > 0) {
    status = 'INVALID';
  } else if (isDuplicate) {
    status = 'DUPLICATE';
  } else if (warnings.length > 0) {
    status = 'WARNING';
  }

  const finalRecord: ForwardingProgressiveRecord = {
    id: mappedRecord.id || (targetModule === 'dispatch' ? `DSP-IMP-${Date.now().toString().slice(-6)}-${rowIndex}` : `FPR-IMP-${Date.now().toString().slice(-6)}-${rowIndex}`),
    month: mappedRecord.month || 'August 2026',
    coordinator,
    client: clientName,
    clientId,
    modeOfShipment,
    area,
    referenceNumber: referenceNumber || mappedRecord.referenceNumber || '',
    actualDispatchDate: actualDispatchDate || '',
    expectedDeliveryDate: expectedLeadtimeRes.expectedDeliveryDate || undefined,
    expectedDeliveryDateFormatted: expectedLeadtimeRes.expectedDeliveryDateFormatted || undefined,
    requestDeliveryDate: requestDeliveryDate || undefined,
    leadtimeStatus: expectedLeadtimeRes.status,
    leadtimeMessage: expectedLeadtimeRes.message,
    plannedDeliveryDate: plannedDeliveryDate || undefined,
    destination: mappedRecord.destination || (consignee ? `${consignee} Terminal` : 'Main Destination'),
    destinationCode: mappedRecord.destinationCode || (area === 'NCR' ? 'MNL-01' : (area === 'Visayas' ? 'CEB-01' : 'DVO-01')),
    consignee,
    quantity,
    unit: mappedRecord.unit || 'Boxes',
    courier: mappedRecord.courier || 'OFII Fleet Logistics',
    truckProvider: mappedRecord.truckProvider || mappedRecord.courier || 'OFII Fleet Logistics',
    plateNumber: mappedRecord.plateNumber,
    driverName: mappedRecord.driverName,
    timeArrived: mappedRecord.timeArrived,
    startLoadingTime: mappedRecord.startLoadingTime,
    endLoadingTime: mappedRecord.endLoadingTime,
    actualDepartureTime: mappedRecord.actualDepartureTime,
    truckArrivalTime: mappedRecord.truckArrivalTime || mappedRecord.timeArrived,
    loadingStartTime: mappedRecord.loadingStartTime || mappedRecord.startLoadingTime,
    loadingEndTime: mappedRecord.loadingEndTime || mappedRecord.endLoadingTime,
    departureTime: mappedRecord.departureTime || mappedRecord.actualDepartureTime,
    cbm,
    volumeWeightKg,
    actualWeightKg,
    chargeableWeightFees: mappedRecord.chargeableWeightFees || 'PHP 25,000.00',
    declaredValue: mappedRecord.declaredValue || 'PHP 1,000,000.00',
    podNumber: mappedRecord.podNumber || `POD-${Math.floor(100000 + Math.random() * 900000)}`,
    awbCourierRefNumber: mappedRecord.awbCourierRefNumber || `AWB-${Math.floor(10000 + Math.random() * 90000)}`,
    deliveryStatus: normalizeDeliveryStatus(mappedRecord.deliveryStatus, !!actualDeliveryDate),
    receiversName: mappedRecord.receiversName,
    actualDeliveryDate,
    deliveryLeadTimeDays,
    deliveryTatDays,
    deliveryPerformance,
    reasonForDelay: mappedRecord.reasonForDelay,
    podStatus: autoPodRes.status,
    dateOfPodReturn,
    podLeadTimeDays: podDueRes.podLeadTimeDays,
    podReturnDueDate: podDueRes.podReturnDueDate || undefined,
    podReturnDueDateFormatted: podDueRes.podReturnDueDateFormatted || undefined,
    podTatDays,
    podPerformance,
    podReasonForDelay: mappedRecord.podReasonForDelay,
    isDeleted: false,
    importSource: 'Excel Import',
    importedAt: mappedRecord.importedAt || new Date().toISOString(),
  };

  return {
    rowIndex,
    originalRow,
    mappedRecord: finalRecord,
    status,
    errors,
    warnings,
    cellErrors,
    cellWarnings,
    isDuplicate,
    duplicateReason,
    isClientRecognized,
    recognizedClient,
    assignedCoordinator: coordinator,
    isSelectedForImport: status === 'VALID' || status === 'WARNING',
    duplicateAction,
  };
}

/**
 * Validates all parsed Excel rows against mapping, schema requirements, and business rules.
 */
export function validateImportRows(
  rows: Record<string, any>[],
  columnMapping: Record<string, OFIIFieldKey>,
  clients: ClientSummary[],
  existingForwardingRecords: ForwardingProgressiveRecord[],
  existingDispatches: DispatchRecord[] = [],
  targetModule: 'forwarding' | 'dispatch' = 'forwarding'
): ValidationSummary {
  // Sets of existing identifiers for duplicate detection
  const existingRefs = new Set<string>();
  const existingPods = new Set<string>();
  const existingAwbs = new Set<string>();

  existingForwardingRecords.forEach((r) => {
    if (r.referenceNumber) existingRefs.add(String(r.referenceNumber).trim().toLowerCase());
    if (r.podNumber) existingPods.add(String(r.podNumber).trim().toLowerCase());
    if (r.awbCourierRefNumber) existingAwbs.add(String(r.awbCourierRefNumber).trim().toLowerCase());
  });

  existingDispatches.forEach((d) => {
    if (d.podNumber) existingPods.add(String(d.podNumber).trim().toLowerCase());
  });

  // Track batch internal duplicates
  const batchRefs = new Map<string, number>();
  const batchPods = new Map<string, number>();

  const validatedRows: ValidatedImportRow[] = rows.map((row, index) => {
    // Extract values based on mapping
    const extracted: Partial<Record<OFIIFieldKey, any>> = {};
    Object.entries(columnMapping).forEach(([excelCol, fieldKey]) => {
      if (fieldKey && fieldKey !== 'none') {
        extracted[fieldKey] = row[excelCol];
      }
    });

    const mappedPartial: Partial<ForwardingProgressiveRecord> = {
      client: extracted.client !== undefined ? String(extracted.client).trim() : undefined,
      consignee: extracted.consignee !== undefined ? String(extracted.consignee).trim() : undefined,
      modeOfShipment: extracted.modeOfShipment,
      area: extracted.area,
      referenceNumber: extracted.referenceNumber !== undefined ? String(extracted.referenceNumber).trim() : undefined,
      actualDispatchDate: extracted.actualDispatchDate,
      plannedDeliveryDate: extracted.plannedDeliveryDate,
      requestDeliveryDate: extracted.requestDeliveryDate,
      actualDeliveryDate: extracted.actualDeliveryDate,
      dateOfPodReturn: extracted.dateOfPodReturn,
      quantity: extracted.quantity,
      cbm: extracted.cbm,
      volumeWeightKg: extracted.volumeWeightKg,
      actualWeightKg: extracted.actualWeightKg,
      chargeableWeightFees: extracted.chargeableWeightFees,
      declaredValue: extracted.declaredValue,
      podNumber: extracted.podNumber !== undefined ? String(extracted.podNumber).trim() : undefined,
      awbCourierRefNumber: extracted.awbCourierRefNumber !== undefined ? String(extracted.awbCourierRefNumber).trim() : undefined,
      destination: extracted.destination,
      destinationCode: extracted.destinationCode,
      courier: extracted.courier,
      truckProvider: extracted.truckProvider || extracted.courier,
      plateNumber: extracted.plateNumber,
      driverName: extracted.driverName,
      receiversName: extracted.receiversName,
      unit: extracted.unit,
      month: extracted.month,
      timeArrived: extracted.timeArrived || extracted.truckArrivalTime,
      startLoadingTime: extracted.startLoadingTime || extracted.loadingStartTime,
      endLoadingTime: extracted.endLoadingTime || extracted.loadingEndTime,
      actualDepartureTime: extracted.actualDepartureTime || extracted.departureTime,
      reasonForDelay: extracted.reasonForDelay,
      podReasonForDelay: extracted.podReasonForDelay,
    };

    const rowRes = validateSingleRow({
      rowIndex: index + 1,
      originalRow: row,
      mappedRecord: mappedPartial,
      clients,
      existingRefs,
      existingPods,
      existingAwbs,
      batchRefs,
      batchPods,
      targetModule,
    });

    // Register batch keys
    if (rowRes.mappedRecord.referenceNumber) {
      const refKey = String(rowRes.mappedRecord.referenceNumber).trim().toLowerCase();
      if (!batchRefs.has(refKey)) {
        batchRefs.set(refKey, index);
      }
    }
    if (rowRes.mappedRecord.podNumber) {
      const podKey = String(rowRes.mappedRecord.podNumber).trim().toLowerCase();
      if (!batchPods.has(podKey)) {
        batchPods.set(podKey, index);
      }
    }

    return rowRes;
  });

  const totalRows = validatedRows.length;
  const validCount = validatedRows.filter((r) => r.status === 'VALID').length;
  const warningCount = validatedRows.filter((r) => r.status === 'WARNING').length;
  const invalidCount = validatedRows.filter((r) => r.status === 'INVALID').length;
  const duplicateCount = validatedRows.filter((r) => r.status === 'DUPLICATE').length;
  const readyToImportCount = validatedRows.filter((r) => r.isSelectedForImport && r.status !== 'INVALID').length;

  return {
    totalRows,
    validCount,
    warningCount,
    invalidCount,
    duplicateCount,
    readyToImportCount,
    rows: validatedRows,
  };
}

