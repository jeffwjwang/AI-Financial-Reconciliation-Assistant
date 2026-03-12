import { GoogleGenAI, Type } from "@google/genai";
import { db } from "./db";

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

export const parseFinancePDF = async (base64Data: string, subject?: string) => {
  const model = "gemini-3-flash-preview";
  
  const response = await ai.models.generateContent({
    model: model,
    contents: [
      {
        parts: [
          {
            inlineData: {
              mimeType: "application/pdf",
              data: base64Data,
            },
          },
          {
            text: `You are a financial expert. Extract structured data from this document. 
            Context: The email subject was "${subject || 'Unknown'}".
            If it is a statement, identify the vendor, total amount, period (YYYY-MM), issue date (YYYY-MM-DD), and statement number. 
            Also extract all invoice numbers mentioned in the statement.
            If it is an invoice, identify the vendor, amount, date, and invoice number.
            Also extract a brief summary of the purpose/subject of the document.`,
          },
        ],
      },
    ],
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          type: { type: Type.STRING, description: "Either 'statement' or 'invoice'" },
          vendor: { type: Type.STRING },
          amount: { type: Type.NUMBER },
          date: { type: Type.STRING, description: "Period for statement (YYYY-MM) or date for invoice" },
          issueDate: { type: Type.STRING, nullable: true, description: "Specific date on the document (YYYY-MM-DD)" },
          statementNumber: { type: Type.STRING, nullable: true },
          invoiceNumber: { type: Type.STRING, nullable: true, description: "Only for invoices" },
          invoiceNumbers: { 
            type: Type.ARRAY, 
            items: { type: Type.STRING }, 
            description: "List of invoice numbers mentioned in a statement" 
          },
          items: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                invoiceNumber: { type: Type.STRING },
                amount: { type: Type.NUMBER }
              },
              required: ["invoiceNumber", "amount"]
            },
            description: "Breakdown of invoices and their amounts in the statement"
          },
          summary: { type: Type.STRING },
        },
        required: ["type", "vendor", "amount", "date", "summary"],
      },
    },
  });

  try {
    return JSON.parse(response.text || "{}");
  } catch (e) {
    throw new Error("AI Parsing failed");
  }
};

export const matchInvoices = async () => {
  const pendingStatements = await db.statements.where('status').equals('pending').toArray();
  const allInvoices = await db.invoices.toArray();

  for (const statement of pendingStatements) {
    const match = allInvoices.find(inv => {
      const vendorMatch = inv.vendor.toLowerCase().includes(statement.vendor.toLowerCase()) || 
                         statement.vendor.toLowerCase().includes(inv.vendor.toLowerCase());
      const amountMatch = Math.abs(inv.amount - statement.amount) < 0.01;
      return vendorMatch && amountMatch;
    });

    if (match) {
      await db.statements.update(statement.id!, {
        status: 'matched',
        matchedInvoiceId: match.id
      });
      await db.invoices.update(match.id!, {
        matchedStatementId: statement.id
      });
    }
  }
};

