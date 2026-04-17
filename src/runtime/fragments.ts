import { RuntimeFeature } from './featureDetection';

interface RuntimeFragmentParts {
  declarations: string;
  procedures: string;
}

const fragmentByFeature: Record<RuntimeFeature, string> = {
  'console.log': [
    "' --- vbalib: console.log ---",
    'Public Sub TS_ConsoleLog(ByVal message As Variant)',
    '    Debug.Print CStr(message)',
    'End Sub',
    '',
  ].join('\r\n'),
  'array.push': [
    "' --- vbalib: array.push ---",
    'Public Function TS_ArrayPush(ByRef arr As Variant, ByVal value As Variant) As Long',
    '    Dim currentSize As Long',
    '    If IsEmpty(arr) Then',
    '        ReDim arr(0 To 0)',
    '        If IsObject(value) Then',
    '            Set arr(0) = value',
    '        Else',
    '            arr(0) = value',
    '        End If',
    '        TS_ArrayPush = 1',
    '        Exit Function',
    '    End If',
    '',
    '    currentSize = UBound(arr) - LBound(arr) + 1',
    '    ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)',
    '    If IsObject(value) Then',
    '        Set arr(UBound(arr)) = value',
    '    Else',
    '        arr(UBound(arr)) = value',
    '    End If',
    '    TS_ArrayPush = currentSize + 1',
    'End Function',
    '',
  ].join('\r\n'),
  'error.stack': [
    "' --- vbalib: error.stack ---",
    'Private TS_LastErrorNumber As Long',
    'Private TS_LastErrorDescription As String',
    '',
    'Public Sub TS_PushError(ByVal errorNumber As Long, ByVal errorDescription As String)',
    '    TS_LastErrorNumber = errorNumber',
    '    TS_LastErrorDescription = errorDescription',
    'End Sub',
    '',
    'Public Function TS_LastErrorMessage() As String',
    '    TS_LastErrorMessage = TS_LastErrorDescription',
    'End Function',
    '',
    'Public Sub TS_ClearError()',
    '    TS_LastErrorNumber = 0',
    '    TS_LastErrorDescription = ""',
    'End Sub',
    '',
  ].join('\r\n'),
  'iterator.protocol': [
    "' --- vbalib: iterator.protocol ---",
    'Public Function TS_HasArrayBounds(ByRef arr As Variant) As Boolean',
    '    On Error GoTo TS_ITER_NO_BOUNDS',
    '    Dim lower As Long',
    '    Dim upper As Long',
    '    lower = LBound(arr)',
    '    upper = UBound(arr)',
    '    TS_HasArrayBounds = (upper >= lower)',
    '    Exit Function',
    'TS_ITER_NO_BOUNDS:',
    '    TS_HasArrayBounds = False',
    'End Function',
    '',
  ].join('\r\n'),
};

function splitFragment(fragment: string): RuntimeFragmentParts {
  const lines = fragment.split(/\r\n|\n|\r/);
  const declarations: string[] = [];
  const procedures: string[] = [];
  let inProcedure = false;

  for (const line of lines) {
    const trimmed = line.trim();

    if (/^(Public\s+(Sub|Function)|Private\s+(Sub|Function))/i.test(trimmed)) {
      inProcedure = true;
      procedures.push(line);
      continue;
    }

    if (/^End\s+(Sub|Function)$/i.test(trimmed)) {
      procedures.push(line);
      inProcedure = false;
      continue;
    }

    if (inProcedure) {
      procedures.push(line);
      continue;
    }

    if (/^(Private|Public\s+Const)/i.test(trimmed)) {
      declarations.push(line);
    } else if (trimmed.length > 0) {
      procedures.push(line);
    }
  }

  return {
    declarations: declarations.join('\r\n'),
    procedures: procedures.join('\r\n'),
  };
}

export function renderRuntimeSections(features: Set<RuntimeFeature>): RuntimeFragmentParts {
  const sorted = Array.from(features).sort();
  if (!sorted.length) {
    return { declarations: '', procedures: '' };
  }

  const declarationChunks: string[] = [];
  const procedureChunks: string[] = [];

  for (const feature of sorted) {
    const parts = splitFragment(fragmentByFeature[feature]);
    if (parts.declarations) {
      declarationChunks.push(parts.declarations);
    }
    if (parts.procedures) {
      procedureChunks.push(parts.procedures);
    }
  }

  return {
    declarations: declarationChunks.join('\r\n'),
    procedures: procedureChunks.join('\r\n'),
  };
}

export function renderRuntimeFragments(features: Set<RuntimeFeature>): string {
  const sections = renderRuntimeSections(features);
  if (!sections.declarations && !sections.procedures) {
    return '';
  }

  const chunks = [sections.declarations, sections.procedures].filter(Boolean);
  return ["' ===== TSTVBA Runtime (auto-injected) =====", ...chunks].join('\r\n');
}
