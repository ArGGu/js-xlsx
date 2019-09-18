export namespace CFB {
  function parse(file: any): any;
  function read(blob: any, options: any): any;
  namespace utils {
    function CheckField(hexstr: any, fld: any): void;
    class ReadShift {
      constructor(size: any, t: any);
      l: any;
    }
    function bconcat(bufs: any): any;
    const consts: {
      DIFSECT: number;
      ENDOFCHAIN: number;
      EntryTypes: any[];
      FATSECT: number;
      FREESECT: number;
      HEADER_CLSID: string;
      HEADER_MINOR_VERSION: string;
      HEADER_SIGNATURE: string;
      MAXREGSECT: number;
      MAXREGSID: number;
      NOSTREAM: number;
    };
    function prep_blob(blob: any, pos: any): void;
  }
  const version: string;
}
export namespace SSF {
  function format(fmt: any, v: any, o: any): any;
  function get_table(): any;
  function load(fmt: any, idx: any): void;
  function load_table(tbl: any): void;
  const opts: string[][];
  function parse_date_code(v: any, opts: any, b2: any): any;
  const version: string;
}
export function parse_xlscfb(cfb: any, options: any): any;
export function parse_zip(zip: any, opts: any): any;
export function read(data: any, opts: any): any;
export function readFile(data: any, opts: any): any;
export function readFileSync(data: any, opts: any): any;
export namespace utils {
  function decode_cell(cstr: any): any;
  function decode_col(colstr: any): any;
  function decode_range(range: any): any;
  function decode_row(rowstr: any): any;
  function encode_cell(cell: any): any;
  function encode_col(col: any): any;
  function encode_range(cs: any, ce: any): any;
  function encode_row(row: any): any;
  function format_cell(cell: any, v: any): any;
  function get_formulae(sheet: any): any;
  function make_csv(sheet: any, opts: any): any;
  function make_formulae(sheet: any): any;
  function make_json(sheet: any, opts: any): any;
  function sheet_to_csv(sheet: any, opts: any): any;
  function sheet_to_formulae(sheet: any): any;
  function sheet_to_json(sheet: any, opts: any): any;
  function sheet_to_row_object_array(sheet: any, opts: any): any;
  function split_cell(cstr: any): any;
}
export const version: string;
export function write(wb: any, opts: any): any;
export function writeFile(wb: any, filename: any, opts: any): any;
export function writeFileSync(wb: any, filename: any, opts: any): any;
