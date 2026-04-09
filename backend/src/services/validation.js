function createError(message, statusCode = 400) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function normalizeListValue(value) {
  if (value === undefined || value === null || value === "") return [];
  if (Array.isArray(value)) return value.map((item) => String(item).trim()).filter(Boolean);
  return String(value)
    .split(/[\n,]+/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function validateField(field, rawValue) {
  if (field.type === "checkbox") {
    if (rawValue === undefined) {
      return Boolean(field.defaultValue);
    }
    if (typeof rawValue !== "boolean") {
      throw createError(`Field '${field.label}' must be true or false.`);
    }
    return rawValue;
  }

  if (field.type === "multiselect") {
    const values = normalizeListValue(rawValue);
    if (field.required && values.length === 0) {
      throw createError(`Field '${field.label}' requires at least one selection.`);
    }

    if (field.options?.length) {
      const invalid = values.filter((value) => !field.options.includes(value));
      if (invalid.length > 0) {
        throw createError(`Field '${field.label}' contains invalid options: ${invalid.join(", ")}.`);
      }
    }

    return values;
  }

  if (rawValue === undefined || rawValue === null || rawValue === "") {
    if (field.required) {
      throw createError(`Field '${field.label}' is required.`);
    }
    return field.defaultValue ?? "";
  }

  if (field.type === "number") {
    const numberValue = typeof rawValue === "number" ? rawValue : Number(rawValue);
    if (!Number.isFinite(numberValue)) {
      throw createError(`Field '${field.label}' must be a valid number.`);
    }
    if (field.min !== undefined && numberValue < field.min) {
      throw createError(`Field '${field.label}' must be at least ${field.min}.`);
    }
    if (field.max !== undefined && numberValue > field.max) {
      throw createError(`Field '${field.label}' must be at most ${field.max}.`);
    }
    return numberValue;
  }

  if (typeof rawValue !== "string") {
    throw createError(`Field '${field.label}' must be text.`);
  }

  const trimmed = rawValue.trim();
  if (field.required && !trimmed) {
    throw createError(`Field '${field.label}' is required.`);
  }

  if (field.options?.length && trimmed && !field.options.includes(trimmed)) {
    throw createError(`Field '${field.label}' must be one of: ${field.options.join(", ")}.`);
  }

  return trimmed;
}

export function validatePayload(script, payload = {}) {
  const allowedFields = new Set((script.fields || []).map((field) => field.id));

  for (const key of Object.keys(payload)) {
    if (!allowedFields.has(key)) {
      throw createError(`Unexpected field '${key}' was provided.`);
    }
  }

  return (script.fields || []).reduce((acc, field) => {
    acc[field.id] = validateField(field, payload[field.id]);
    return acc;
  }, {});
}

export { createError, normalizeListValue };
