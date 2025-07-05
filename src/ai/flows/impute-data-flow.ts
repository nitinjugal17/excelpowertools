
'use server';

/**
 * @fileOverview An AI flow for imputing missing data in a spreadsheet.
 *
 * - imputeData - A function that suggests values for multiple blank cells based on context.
 * - ImputeDataInput - The input type for the imputeData function.
 * - ImputeDataOutput - The return type for the imputeData function.
 */

import { ai } from '@/ai/genkit';
import { z } from 'genkit';

const ImputationTargetSchema = z.object({
  identifier: z.string().describe('A unique identifier for this row, like a cell address (e.g., "A2").'),
  rowData: z.record(z.any()).describe('The key-value map of the row data that has a blank cell.'),
});

const ImputeDataInputSchema = z.object({
  headers: z.array(z.string()).describe('The headers of the spreadsheet.'),
  targetColumn: z.string().describe('The name of the column that has blank cells to be filled.'),
  rowsToImpute: z.array(ImputationTargetSchema).describe('An array of rows that need imputation.'),
  exampleRows: z.array(z.record(z.any())).describe('An array of other complete rows from the sheet to provide context and show data patterns.'),
});
export type ImputeDataInput = z.infer<typeof ImputeDataInputSchema>;

const ImputationSuggestionSchema = z.object({
  identifier: z.string().describe('The unique identifier for the row this suggestion belongs to.'),
  suggestion: z.string().describe('The suggested value for the blank cell. This should be the value only, without any explanation or extra text. If no suggestion can be made, return an empty string.'),
});

const ImputeDataOutputSchema = z.object({
  suggestions: z.array(ImputationSuggestionSchema).describe('An array of suggestions for the imputed cells.'),
});
export type ImputeDataOutput = z.infer<typeof ImputeDataOutputSchema>;


export async function imputeData(input: ImputeDataInput): Promise<ImputeDataOutput> {
  return imputeDataFlow(input);
}

const imputeDataPrompt = ai.definePrompt({
  name: 'imputeDataPrompt',
  input: { schema: ImputeDataInputSchema },
  output: { schema: ImputeDataOutputSchema },
  prompt: `You are a data analysis expert. Your task is to infer the most logical value for blank cells in a list of spreadsheet rows.
I will provide you with the headers of the spreadsheet, the name of the column that is blank ('{{targetColumn}}'), a list of rows that need a value for that column, and several example rows that are complete to show the data patterns.

For each row in the 'rowsToImpute' list, determine the most likely value for the '{{targetColumn}}' column.
Return a list of suggestions. Each suggestion object must include the original 'identifier' from the input row and the 'suggestion' string.
Return ONLY the suggested value, with no extra explanation, preamble, or formatting. If you cannot determine a value for a specific row, return an empty string for its suggestion.

Spreadsheet Headers:
{{json headers}}

Example Complete Rows (for context):
{{#each exampleRows}}
- {{json this}}
{{/each}}

Rows to Impute (each has a blank in the '{{targetColumn}}' column):
{{#each rowsToImpute}}
- Identifier: {{this.identifier}}
  Row Data: {{json this.rowData}}
{{/each}}

Based on the patterns from the example rows, what are the most likely values for the '{{targetColumn}}' column in each of the "Rows to Impute"?
`,
});

const imputeDataFlow = ai.defineFlow(
  {
    name: 'imputeDataFlow',
    inputSchema: ImputeDataInputSchema,
    outputSchema: ImputeDataOutputSchema,
  },
  async (input) => {
    if (input.rowsToImpute.length === 0) {
        return { suggestions: [] };
    }
    const { output } = await imputeDataPrompt(input);
    return output || { suggestions: [] };
  }
);
