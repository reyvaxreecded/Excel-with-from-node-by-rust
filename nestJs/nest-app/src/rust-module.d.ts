declare module 'excel-manager-native' {
  export function read_excel_file(fileName: string): string[][];
  export function upsert_row(
    fileName: string,
    sheetName: string,
    row: UpdateCell[],
  ): void;

  export interface UpdateCell {
    cell: number;
    value: string;
  }
}
