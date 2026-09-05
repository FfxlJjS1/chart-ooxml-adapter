/** Structured errors thrown by the adapter. */

export type AdapterErrorCode =
  | "INVALID_INPUT"
  | "UNSUPPORTED_SOURCE"
  | "UNSUPPORTED_CHART_TYPE"
  | "INVALID_UNIVERSAL_MODEL"
  | "OFFICE_OPEN_VALIDATION";

export class ChartAdapterError extends Error {
  readonly code: AdapterErrorCode;
  readonly details?: readonly string[];

  constructor(code: AdapterErrorCode, message: string, details?: readonly string[]) {
    super(message);
    this.name = "ChartAdapterError";
    this.code = code;
    this.details = details;
  }
}

export class InvalidChartInputError extends ChartAdapterError {
  constructor(message: string, details?: readonly string[]) {
    super("INVALID_INPUT", message, details);
    this.name = "InvalidChartInputError";
  }
}

export class UnsupportedChartSourceError extends ChartAdapterError {
  constructor(message: string, details?: readonly string[]) {
    super("UNSUPPORTED_SOURCE", message, details);
    this.name = "UnsupportedChartSourceError";
  }
}

export class UnsupportedChartTypeError extends ChartAdapterError {
  constructor(message: string, details?: readonly string[]) {
    super("UNSUPPORTED_CHART_TYPE", message, details);
    this.name = "UnsupportedChartTypeError";
  }
}

export class InvalidUniversalModelError extends ChartAdapterError {
  constructor(message: string, details?: readonly string[]) {
    super("INVALID_UNIVERSAL_MODEL", message, details);
    this.name = "InvalidUniversalModelError";
  }
}

export class OfficeOpenValidationError extends ChartAdapterError {
  constructor(message: string, details?: readonly string[]) {
    super("OFFICE_OPEN_VALIDATION", message, details);
    this.name = "OfficeOpenValidationError";
  }
}
