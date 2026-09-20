import { createMcpHandler } from '@vercel/mcp-adapter';
import DocumentIntelligence, {
  isUnexpected,
  getLongRunningPoller,
  AnalyzeOperationOutput
} from '@azure-rest/ai-document-intelligence';
import { z } from 'zod';

const endpoint = process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT || '';
const apiKey = process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY || '';

// Initialize client
const client = DocumentIntelligence(endpoint, { key: apiKey });

const handler = createMcpHandler((server) => {
  // -------------------------------------------------------------
  // Tool 1: Analyze Invoice (Base64 or URL)
  // -------------------------------------------------------------
  server.tool(
    'analyze_invoice',
    'Extracts key data fields (Vendor, Total, Dates, Line Items) from an invoice PDF or image.',
    {
      base64Data: z.string().optional().describe('Base64 encoded string of invoice file (PDF, JPG, PNG)'),
      urlSource: z.string().url().optional().describe('Public HTTP/HTTPS URL of the invoice')
    },
    async ({ base64Data, urlSource }) => {
      try {
        if (!base64Data && !urlSource) {
          throw new Error('Either base64Data or urlSource must be provided.');
        }

        const bodyPayload = urlSource
          ? { urlSource }
          : { base64Source: base64Data!.replace(/^data:.*?;base64,/, '') };

        const initialResponse = await client
          .path('/documentModels/{modelId}:analyze', 'prebuilt-invoice')
          .post({
            contentType: 'application/json',
            body: bodyPayload
          });

        if (isUnexpected(initialResponse)) {
          throw new Error(initialResponse.body.error?.message || 'Invoice analysis failed.');
        }

        const poller = getLongRunningPoller(client, initialResponse);
        const result = (await poller.pollUntilDone()).body as AnalyzeOperationOutput;
        const analyzeResult = result.analyzeResult;

        const document = analyzeResult?.documents?.[0];
        if (!document) {
          return { content: [{ type: 'text', text: 'No structured invoice data detected.' }] };
        }

        const fields = document.fields || {};
        const extractedSummary = [
          `Vendor Name: ${fields.VendorName?.content || 'N/A'}`,
          `Customer Name: ${fields.CustomerName?.content || 'N/A'}`,
          `Invoice ID: ${fields.InvoiceId?.content || 'N/A'}`,
          `Invoice Date: ${fields.InvoiceDate?.content || 'N/A'}`,
          `Due Date: ${fields.DueDate?.content || 'N/A'}`,
          `SubTotal: ${fields.SubTotal?.content || 'N/A'}`,
          `Total Tax: ${fields.TotalTax?.content || 'N/A'}`,
          `Invoice Total: ${fields.InvoiceTotal?.content || 'N/A'}`
        ].join('\n');

        return { content: [{ type: 'text', text: extractedSummary }] };
      } catch (error: any) {
        return {
          content: [{ type: 'text', text: `Error processing invoice: ${error.message}` }],
          isError: true
        };
      }
    }
  );

  // -------------------------------------------------------------
  // Tool 2: Analyze Layout & Extract Tables (Base64 or URL)
  // -------------------------------------------------------------
  server.tool(
    'analyze_layout_tables',
    'Extracts layout structure, paragraphs, and tables from complex documents or PDFs.',
    {
      base64Data: z.string().optional().describe('Base64 encoded string of document file'),
      urlSource: z.string().url().optional().describe('Public URL of the document')
    },
    async ({ base64Data, urlSource }) => {
      try {
        if (!base64Data && !urlSource) {
          throw new Error('Either base64Data or urlSource must be provided.');
        }

        const bodyPayload = urlSource
          ? { urlSource }
          : { base64Source: base64Data!.replace(/^data:.*?;base64,/, '') };

        const initialResponse = await client
          .path('/documentModels/{modelId}:analyze', 'prebuilt-layout')
          .post({
            contentType: 'application/json',
            body: bodyPayload
          });

        if (isUnexpected(initialResponse)) {
          throw new Error(initialResponse.body.error?.message || 'Layout analysis failed.');
        }

        const poller = getLongRunningPoller(client, initialResponse);
        const result = (await poller.pollUntilDone()).body as AnalyzeOperationOutput;
        const analyzeResult = result.analyzeResult;

        const tablesCount = analyzeResult?.tables?.length || 0;
        const paragraphs = analyzeResult?.paragraphs?.map((p) => p.content).join('\n') || '';

        const output = [
          `Total Pages: ${analyzeResult?.pages?.length || 0}`,
          `Total Tables Extracted: ${tablesCount}`,
          `\n--- Text Content ---`,
          paragraphs.length > 0 ? paragraphs : 'No paragraph text extracted.'
        ].join('\n');

        return { content: [{ type: 'text', text: output }] };
      } catch (error: any) {
        return {
          content: [{ type: 'text', text: `Error reading layout: ${error.message}` }],
          isError: true
        };
      }
    }
  );

  // -------------------------------------------------------------
  // Tool 3: Analyze ID / Passport Documents
  // -------------------------------------------------------------
  server.tool(
    'analyze_id_document',
    'Extracts identity information from passports, driver licenses, or ID cards.',
    {
      base64Data: z.string().optional().describe('Base64 encoded image of ID document'),
      urlSource: z.string().url().optional().describe('Public URL of ID image')
    },
    async ({ base64Data, urlSource }) => {
      try {
        if (!base64Data && !urlSource) {
          throw new Error('Either base64Data or urlSource must be provided.');
        }

        const bodyPayload = urlSource
          ? { urlSource }
          : { base64Source: base64Data!.replace(/^data:.*?;base64,/, '') };

        const initialResponse = await client
          .path('/documentModels/{modelId}:analyze', 'prebuilt-idDocument')
          .post({
            contentType: 'application/json',
            body: bodyPayload
          });

        if (isUnexpected(initialResponse)) {
          throw new Error(initialResponse.body.error?.message || 'ID analysis failed.');
        }

        const poller = getLongRunningPoller(client, initialResponse);
        const result = (await poller.pollUntilDone()).body as AnalyzeOperationOutput;
        const analyzeResult = result.analyzeResult;

        const document = analyzeResult?.documents?.[0];
        if (!document) {
          return { content: [{ type: 'text', text: 'No structured ID document data detected.' }] };
        }

        const fields = document.fields || {};
        const extractedSummary = [
          `First Name: ${fields.FirstName?.content || 'N/A'}`,
          `Last Name: ${fields.LastName?.content || 'N/A'}`,
          `Document Number: ${fields.DocumentNumber?.content || 'N/A'}`,
          `Date of Birth: ${fields.DateOfBirth?.content || 'N/A'}`,
          `Date of Expiration: ${fields.DateOfExpiration?.content || 'N/A'}`,
          `Country/Region: ${fields.CountryRegion?.content || 'N/A'}`
        ].join('\n');

        return { content: [{ type: 'text', text: extractedSummary }] };
      } catch (error: any) {
        return {
          content: [{ type: 'text', text: `Error processing ID document: ${error.message}` }],
          isError: true
        };
      }
    }
  );
});

export const GET = handler;
export const POST = handler;
