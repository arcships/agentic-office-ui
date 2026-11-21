import type { SchemaBuilderSchema } from "@arcships/vue-ui"

/**
 * Demo fixture for `/schema-builder` (#179): an invoice extraction schema with
 * a nested object, an array of objects, a scalar array and a required field.
 */
export const SCHEMA_BUILDER_FIXTURE: SchemaBuilderSchema = {
  properties: [
    {
      id: "field-vendor",
      name: "vendor_name",
      type: "string",
      required: true,
      description: "Legal name of the supplier on the invoice.",
    },
    {
      id: "field-total",
      name: "invoice_total",
      type: "number",
      required: true,
      description: "Grand total in the invoice currency.",
    },
    {
      id: "field-paid",
      name: "is_paid",
      type: "boolean",
      required: false,
      description: "Whether the invoice shows a paid stamp.",
    },
    {
      id: "field-remit",
      name: "remit_to",
      type: "object",
      required: false,
      description: "Remittance address block.",
      properties: [
        { id: "field-remit-street", name: "street", type: "string", required: true, description: "" },
        { id: "field-remit-city", name: "city", type: "string", required: true, description: "" },
        { id: "field-remit-country", name: "country", type: "string", required: false, description: "" },
      ],
    },
    {
      id: "field-lines",
      name: "line_items",
      type: "array",
      itemType: "object",
      required: true,
      description: "One entry per invoice line.",
      properties: [
        { id: "field-lines-sku", name: "sku", type: "string", required: true, description: "" },
        { id: "field-lines-qty", name: "quantity", type: "number", required: true, description: "" },
        { id: "field-lines-amount", name: "amount", type: "number", required: false, description: "" },
      ],
    },
    {
      id: "field-tags",
      name: "tags",
      type: "array",
      itemType: "string",
      required: false,
      description: "Free-form routing tags.",
    },
  ],
}

/** A JSON Schema document the demo can import on demand. */
export const SCHEMA_BUILDER_IMPORT_SAMPLE = JSON.stringify(
  {
    type: "object",
    required: ["purchase_order"],
    properties: {
      purchase_order: { type: "string", description: "PO number printed on the invoice." },
      terms: {
        type: "object",
        properties: {
          net_days: { type: "integer" },
          discount_percent: { type: "number" },
        },
        required: ["net_days"],
      },
      approvals: {
        type: "array",
        items: {
          type: "object",
          properties: {
            approver: { type: "string" },
            approved_at: { type: "string" },
          },
        },
      },
    },
  },
  null,
  2,
)
