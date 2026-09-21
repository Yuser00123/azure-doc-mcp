import { createMcpHandler } from '@vercel/mcp-adapter';
import DocumentIntelligence, {
  isUnexpected,
  getLongRunningPoller,
  AnalyzeOperationOutput,
} from '@azure-rest/ai-document-intelligence';
import { z } from 'zod';

// ---------------------------------------------------------------------------
// Runtime config
// ---------------------------------------------------------------------------
export const runtime = 'nodejs';

// 🔴 CRITICAL. Vercel Hobby defaults to maxDuration = 10s. Document
// Intelligence is an async long-running operation — you POST, then poll until
// done. A multi-page PDF routinely takes 15-40s. Without this export the
// function was being killed mid-poll.
export const maxDuration = 60;

const endpoint = process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT;
const apiKey = process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY;
if (!endpoint || !apiKey) {
  throw new Error(
    'AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT and AZURE_DOCUMENT_INTELLIGENCE_KEY must be set'
  );
}

const AUTH_TOKEN = process.env.MCP_AUTH_TOKEN;

// prebuilt-idDocument extracts names, DOB and document numbers. Behind a
// public URL that is a serious exposure, so it is opt-in.
const ENABLE_ID_DOCUMENT = process.env.ENABLE_ID_DOCUMENT === 'true';

const MAX_DOC_BYTES = 3 * 1024 * 1024; // see note: Vercel body limit, not Azure's
const MAX_TEXT_OUT = 8000;             // paragraph dump can be huge
const POLL_DEADLINE_MS = 50_000;       // stay under maxDuration

const client = DocumentIntelligence(endpoint, { key: apiKey });

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function truncate(s: string, n = MAX_TEXT_OUT): string {
  return s.length > n ? `${s.slice(0, n)}\n…[truncated ${s.length - n} chars]` : s;
}

const err = (msg: string) => ({
  content: [{ type: 'text' as const, text: msg }],
  isError: true,
});

function buildPayload(base64Data?: string, urlSource?: string) {
  if (urlSource) return { urlSource };
  if (!base64Data) throw new Error('Either base64Data or urlSource must be provided.');
  const clean = base64Data.replace(/^data:[^;]*;base64,/, '');
  const approxBytes = Math.floor((clean.length * 3) / 4);
  if (approxBytes > MAX_DOC_BYTES) {
    throw new Error(
      `Document is ~${(approxBytes / 1048576).toFixed(1)}MB. Limit is 3MB: base64 inflates ` +
        `size ~33% and Vercel 413s bodies over 4.5MB before this handler runs. ` +
        `Upload it somewhere public and pass urlSource instead.`
    );
  }
  return { base64Source: clean };
}

/** POST to a prebuilt model and poll with a hard deadline. */
async function analyze(modelId: string, body: Record<string, unknown>) {
  const initial = await client
    .path('/documentModels/{modelId}:analyze', modelId)
    .post({ contentType: 'application/json', body });

  if (isUnexpected(initial)) {
    throw new Error(initial.body.error?.message || `${modelId} analysis failed.`);
  }

  const poller = getLongRunningPoller(client, initial);

  // Previously: await poller.pollUntilDone() with no bound. On a slow or large
  // document that ran past the platform timeout and returned a 504.
  const done = await Promise.race([
    poller.pollUntilDone(),
    new Promise<null>((resolve) => setTimeout(() => resolve(null), POLL_DEADLINE_MS)),
  ]);
  if (!done) {
    throw new Error(
      `Analysis did not finish within ${POLL_DEADLINE_MS / 1000}s. ` +
        `Retry with fewer pages or a smaller file.`
    );
  }

  return (done as { body: AnalyzeOperationOutput }).body.analyzeResult;
}

const sourceSchema = {
  base64Data: z.string().optional().describe('Base64 of the document (PDF, JPG, PNG)'),
  urlSource: z
    .string()
    .url()
    .optional()
    .describe('Public URL of the document. PREFERRED over base64Data.'),
};

// ---------------------------------------------------------------------------
// Field renderers
// ---------------------------------------------------------------------------
function renderFields(fields: Record<string, any>, keys: string[]): string {
  return keys
    .map((k) => `${k.replace(/([A-Z])/g, ' $1').trim()}: ${fields[k]?.content ?? 'N/A'}`)
    .join('\n');
}

/** 🔴 Previously tables were COUNTED and then thrown away — the tool is named
 *  analyze_layout_tables but never returned a single table cell. */
function renderTables(tables: any[]): string {
  if (!tables?.length) return 'No tables found.';
  return tables
    .slice(0, 5)
    .map((t, i) => {
      const rows: string[][] = [];
      for (const cell of t.cells || []) {
        rows[cell.rowIndex] = rows[cell.rowIndex] || [];
        rows[cell.rowIndex][cell.columnIndex] = cell.content ?? '';
      }
      const md = rows.filter(Boolean).map((r) => `| ${(r || []).map((c) => c ?? '').join(' | ')} |`);
      return `Table ${i + 1} (${t.rowCount}x${t.columnCount}):\n${md.join('\n')}`;
    })
    .join('\n\n');
}

// ---------------------------------------------------------------------------
// MCP server
// ---------------------------------------------------------------------------
const handler = createMcpHandler(
  (server) => {
    server.tool(
      'analyze_invoice',
      'Extracts structured fields (vendor, customer, invoice id, dates, totals) from an invoice PDF or image using Azure Document Intelligence.',
      sourceSchema,
      async ({ base64Data, urlSource }) => {
        try {
          const result = await analyze('prebuilt-invoice', buildPayload(base64Data, urlSource));
          const doc = result?.documents?.[0];
          if (!doc) return { content: [{ type: 'text', text: 'No structured invoice data detected.' }] };
          const text = renderFields(doc.fields || {}, [
            'VendorName',
            'CustomerName',
            'InvoiceId',
            'InvoiceDate',
            'DueDate',
            'SubTotal',
            'TotalTax',
            'InvoiceTotal',
          ]);
          return { content: [{ type: 'text', text }] };
        } catch (e: any) {
          return err(`Error processing invoice: ${e.message}`);
        }
      }
    );

    server.tool(
      'analyze_layout_tables',
      'Extracts paragraphs, page count and table contents from complex documents or PDFs. Returns tables as markdown.',
      sourceSchema,
      async ({ base64Data, urlSource }) => {
        try {
          const result = await analyze('prebuilt-layout', buildPayload(base64Data, urlSource));
          const paragraphs = result?.paragraphs?.map((p: any) => p.content).join('\n') || '';
          const output = [
            `Total Pages: ${result?.pages?.length || 0}`,
            `Total Tables Extracted: ${result?.tables?.length || 0}`,
            `\n--- Text Content ---`,
            paragraphs ? truncate(paragraphs) : 'No paragraph text extracted.',
            `\n--- Tables ---`,
            truncate(renderTables(result?.tables || []), 4000),
          ].join('\n');
          return { content: [{ type: 'text', text: output }] };
        } catch (e: any) {
          return err(`Error reading layout: ${e.message}`);
        }
      }
    );

    // Opt-in only: extracts PII (name, DOB, document number).
    if (ENABLE_ID_DOCUMENT) {
      server.tool(
        'analyze_id_document',
        'Extracts identity fields from a passport, driver licence or ID card. Enabled only when ENABLE_ID_DOCUMENT=true.',
        sourceSchema,
        async ({ base64Data, urlSource }) => {
          try {
            const result = await analyze('prebuilt-idDocument', buildPayload(base64Data, urlSource));
            const doc = result?.documents?.[0];
            if (!doc) {
              return { content: [{ type: 'text', text: 'No structured ID document data detected.' }] };
            }
            const text = renderFields(doc.fields || {}, [
              'FirstName',
              'LastName',
              'DocumentNumber',
              'DateOfBirth',
              'DateOfExpiration',
              'CountryRegion',
            ]);
            return { content: [{ type: 'text', text }] };
          } catch (e: any) {
            return err(`Error processing ID document: ${e.message}`);
          }
        }
      );
    }
  },
  // NOTE: 2nd arg is MCP `ServerOptions`, not name/version. The adapter
  // hardcodes serverInfo.name to 'mcp-typescript server on vercel' for every
  // server built with it — so your orchestrator registry must key MCP servers
  // by its own configured id ('azure_doc'), never by the initialize response's
  // name, or the two custom servers will collide.
  {},
  {
    // 🔴 THE BUG FIX — see the vision server for the full explanation.
    // Route lives at /api/mcp, but the adapter defaulted to matching "/mcp",
    // so every request hit its fallback: 404 "Not found".
    basePath: '/api',
    maxDuration: 60,
    verboseLogs: false,
  }
);

async function guarded(request: Request): Promise<Response> {
  if (AUTH_TOKEN) {
    const got = request.headers.get('authorization');
    const want = `Bearer ${AUTH_TOKEN}`;
    if (!got || got.length !== want.length || got !== want) {
      return new Response('Unauthorized', { status: 401 });
    }
  }
  return handler(request);
}

export const GET = guarded;
export const POST = guarded;
