import { useEffect, useState } from "react";
import * as pdfjsLib from "pdfjs-dist";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import "./App.css";

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.mjs",
  import.meta.url
).toString();

type ClassEntry = {
  course: string;
  section?: string;
  day: string;
  start: string;
  end: string;
  room: string;
};

const STORAGE_KEY = "eworganizer-data";

const dayNames: Record<string, string> = {
  A: "Saturday",
  S: "Sunday",
  M: "Monday",
  T: "Tuesday",
  W: "Wednesday",
  R: "Thursday",
  F: "Friday",
};

const days = [
  "Saturday",
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
];

/* ============================================================
   TODAY / NEXT CLASS HELPERS
============================================================ */

const jsDayToName: Record<number, string> = {
  0: "Sunday",
  1: "Monday",
  2: "Tuesday",
  3: "Wednesday",
  4: "Thursday",
  5: "Friday",
  6: "Saturday",
};


function timeToMinutes(
  time: string
): number {
  const match =
    time.match(
      /^(\d{1,2}):(\d{2})(AM|PM)$/i
    );

  if (!match) {
    return 0;
  }

  let hour =
    Number(match[1]);

  const minute =
    Number(match[2]);

  const period =
    match[3].toUpperCase();

  if (
    period === "AM" &&
    hour === 12
  ) {
    hour = 0;
  }

  if (
    period === "PM" &&
    hour !== 12
  ) {
    hour += 12;
  }

  return (
    hour * 60 +
    minute
  );
}


function formatToday(
  date: Date
): string {
  return date.toLocaleDateString(
    "en-US",
    {
      weekday: "long",
      month: "long",
      day: "numeric",
    }
  );
}

/* ============================================================
   COURSE / SCHEDULE PATTERNS
============================================================ */

/*
 * EWU course codes:
 *
 * CSE302
 * CSE325
 * MKT7101
 * POP7203
 *
 * No hardcoded course list.
 */
const courseRegex =
  /\b[A-Z]{2,4}\d{3,4}(?:\s+Lab)?\b/g;


/*
 * Example:
 *
 * M 1:30PM-3:00PM 437
 * MW 4:50PM-6:20PM 110
 * R 11:50AM-1:20PM 550
 */
const scheduleRegex =
  /([ASMTWRF]+)\s+(\d{1,2}:\d{2}\s*(?:AM|PM))-(\d{1,2}:\d{2}\s*(?:AM|PM))\s+([A-Za-z0-9-]+(?:\s+\([^)]*\))?)/gi;


/* ============================================================
   HELPERS
============================================================ */

function cleanCourseName(value: string): string {
  return value
    .replace(/\s+/g, " ")
    .trim();
}


function normalizeTime(value: string): string {
  return value.replace(/\s+/g, "");
}

function extractFullStudentName(
  text: string
): string {
  const normalizedText = text
    .replace(/\s+/g, " ")
    .trim();

  // Format:
  // Name: Fariha Nusrat Khan
  const nameLabelMatch =
    normalizedText.match(
      /\bName\s*:\s*((?:Md\.?\s+)?[A-Za-z][A-Za-z.'-]*(?:\s+[A-Za-z][A-Za-z.'-]*){1,5})\b/i
    );

  if (nameLabelMatch?.[1]) {
    return nameLabelMatch[1]
      .replace(/\s+/g, " ")
      .trim();
  }

  // Format:
  // 2024-3-60-446 Fariha Nusrat Khan Fall-2026
  const idFirstMatch =
    normalizedText.match(
      /\b\d{4}-\d+-\d+-\d+\s+((?:Md\.?\s+)?[A-Za-z][A-Za-z.'-]*(?:\s+[A-Za-z][A-Za-z.'-]*){1,5})\s+(?:Fall|Spring|Summer)-\d{4}\b/i
    );

  if (idFirstMatch?.[1]) {
    return idFirstMatch[1]
      .replace(/\s+/g, " ")
      .trim();
  }

  // Format:
  // Fariha Nusrat Khan 2024-3-60-446 Fall-2026
  const nameFirstMatch =
    normalizedText.match(
      /\b((?:Md\.?\s+)?[A-Za-z][A-Za-z.'-]*(?:\s+[A-Za-z][A-Za-z.'-]*){1,5})\s+\d{4}-\d+-\d+-\d+\s+(?:Fall|Spring|Summer)-\d{4}\b/i
    );

  if (nameFirstMatch?.[1]) {
    return nameFirstMatch[1]
      .replace(/\s+/g, " ")
      .trim();
  }

  return "Student";
}

function extractStudentId(
  text: string
): string {
  const match = text.match(
    /\b\d{4}-\d+-\d+-\d+\b/
  );

  return match?.[0] || "";
}

function extractStudentName(
  text: string
): string {
  /*
   * Actual PDF.js output from the EWU slip:
   *
   * 2024-3-60-162 Md. Mahir Hasan Shuvo Fall-2026
   *
   * So the order is:
   *
   * STUDENT ID → NAME → SEMESTER
   */

  const normalizedText = text
    .replace(/\s+/g, " ")
    .trim();

  const match = normalizedText.match(
    /\b\d{4}-\d+-\d+-\d+\s+((?:Md\.?\s+)?[A-Za-z][A-Za-z.'-]*(?:\s+[A-Za-z][A-Za-z.'-]*){1,4})\s+(?:Fall|Spring|Summer)-\d{4}\b/i
  );

  if (!match) {
    return "Student";
  }

  const fullName = match[1]
    .replace(/\s+/g, " ")
    .trim();

  const nameParts = fullName.split(" ");

  // Skip "Md." or "Md" prefix.
  if (
    nameParts.length > 1 &&
    /^md\.?$/i.test(nameParts[0])
  ) {
    return nameParts[1];
  }

  return nameParts[0] || "Student";
}



function extractStudentInfoFromExcel(
  workbook: XLSX.WorkBook
): {
  firstName: string;
  fullName: string;
  studentId: string;
} {
  let fullName = "";
  let studentId = "";

  for (
    const sheetName of workbook.SheetNames
  ) {
    const worksheet =
      workbook.Sheets[sheetName];

    const rows =
      XLSX.utils.sheet_to_json<
        Record<string, unknown>
      >(worksheet, {
        header: 1,
        defval: "",
      });

    for (const row of rows) {
      const values =
        Object.values(row)
          .map((value) =>
            String(value).trim()
          )
          .filter(Boolean);

      for (
        let i = 0;
        i < values.length;
        i++
      ) {
        const value = values[i];

        // Name: Md. Mahir Hasan Shuvo
        if (
          /^name\s*:/i.test(value)
        ) {
          const name =
            value
              .replace(
                /^name\s*:\s*/i,
                ""
              )
              .trim();

          if (name) {
            fullName = name;
          } else if (
            values[i + 1]
          ) {
            fullName =
              values[i + 1];
          }
        }

        // Name | Md. Mahir Hasan Shuvo
        if (
          /^name$/i.test(value) &&
          values[i + 1]
        ) {
          fullName =
            values[i + 1];
        }

        // ID# | 2024-3-60-162
        if (
          /^id#?$/i.test(value) &&
          values[i + 1]
        ) {
          studentId =
            values[i + 1];
        }

        // Direct ID detection
        if (
          /^\d{4}-\d+-\d+-\d+$/.test(
            value
          )
        ) {
          studentId = value;
        }
      }
    }
  }

  fullName = fullName
    .replace(/\s+/g, " ")
    .trim();

  const nameParts =
    fullName.split(" ");

  let firstName =
    nameParts[0] || "Student";

  // Skip Md. / Md
  if (
    nameParts.length > 1 &&
    /^md\.?$/i.test(
      nameParts[0]
    )
  ) {
    firstName =
      nameParts[1];
  }

  return {
    firstName,
    fullName:
      fullName || "Student",
    studentId,
  };
}

/*
 * Find the section number belonging to a course.
 *
 * EWU normally has something similar to:
 *
 * CSE302 6 4.50 27000.00
 * CSE325 10 4.00 24000.00
 * POP7203 1 3.00 18000.00
 *
 * We intentionally search the whole table rather than relying
 * on the exact visual layout of the PDF.
 */
function buildPdfSectionMap(
  tableText: string
): Map<string, string> {
  const sectionMap = new Map<string, string>();

  /*
   * First attempt:
   *
   * COURSE SECTION CREDIT AMOUNT
   */
  const standardRegex =
    /\b([A-Z]{2,4}\d{3,4})\s+(\d{1,3})\s+(\d+(?:\.\d+)?)\s+[\d,.]+/gi;

  for (const match of tableText.matchAll(
    standardRegex
  )) {
    const course = match[1].toUpperCase();
    const section = match[2];

    sectionMap.set(course, section);
  }


  /*
   * Second, more flexible attempt.
   *
   * Some PDF text extraction can slightly alter the spacing.
   *
   * Example:
   *
   * CSE302 6 4.50
   * POP7203 1 3.00
   *
   * We only need the course → section relationship.
   */
  const flexibleRegex =
    /\b([A-Z]{2,4}\d{3,4})\s+(\d{1,3})\s+(?=\d+(?:\.\d+)?\b)/gi;

  for (const match of tableText.matchAll(
    flexibleRegex
  )) {
    const course = match[1].toUpperCase();
    const section = match[2];

    if (!sectionMap.has(course)) {
      sectionMap.set(course, section);
    }
  }

  return sectionMap;
}


/* ============================================================
   SCHEDULE PARSER
============================================================ */

function expandScheduleText(
  text: string,
  course: string,
  section?: string
): ClassEntry[] {
  const results: ClassEntry[] = [];

  /*
   * Reset regex state before using matchAll.
   */
  scheduleRegex.lastIndex = 0;

  const matches = [
    ...text.matchAll(scheduleRegex),
  ];

  for (const match of matches) {
    const dayCodes =
      match[1].toUpperCase();

    const start =
      normalizeTime(match[2]);

    const end =
      normalizeTime(match[3]);

    const room =
      match[4].trim();


    /*
     * Expand:
     *
     * MW → Monday + Wednesday
     * S  → Sunday
     * R  → Thursday
     */
    for (const code of dayCodes) {
      if (!dayNames[code]) {
        continue;
      }

      results.push({
        course,
        section,
        day: dayNames[code],
        start,
        end,
        room,
      });
    }
  }

  return results;
}


/* ============================================================
   REMOVE DUPLICATES
============================================================ */

function removeDuplicates(
  entries: ClassEntry[]
): ClassEntry[] {
  return entries.filter(
    (item, index, array) =>
      index ===
      array.findIndex(
        (other) =>
          other.course === item.course &&
          other.section === item.section &&
          other.day === item.day &&
          other.start === item.start &&
          other.end === item.end &&
          other.room === item.room
      )
  );
}


/* ============================================================
   PDF PARSER
============================================================ */

function parsePdfText(
  text: string
): ClassEntry[] {
  const results: ClassEntry[] = [];


  /*
   * Find the beginning of the course table.
   */
  const tableStartMatch =
    text.match(/Course\(s\)/i);

  if (
    !tableStartMatch ||
    tableStartMatch.index === undefined
  ) {
    return [];
  }


  const tableStart =
    tableStartMatch.index;

  const afterTableStart =
    text.substring(tableStart);


  /*
   * Stop before tuition/bank information.
   */
  const tuitionIndex =
    afterTableStart.search(
      /Tuition Fee/i
    );


  const tableText =
    tuitionIndex >= 0
      ? afterTableStart.substring(
          0,
          tuitionIndex
        )
      : afterTableStart;


  /*
   * Find every course automatically.
   */
  courseRegex.lastIndex = 0;

  const courseMatches = [
    ...tableText.matchAll(courseRegex),
  ];


  if (courseMatches.length === 0) {
    return [];
  }


  /*
   * IMPORTANT:
   *
   * Build section information BEFORE parsing
   * individual course blocks.
   *
   * This prevents PDF text-layout differences from
   * causing sections to disappear.
   */
  const sectionMap =
    buildPdfSectionMap(tableText);


  for (
    let i = 0;
    i < courseMatches.length;
    i++
  ) {
    const match =
      courseMatches[i];


    const rawCourse =
      match[0];


    const course =
      cleanCourseName(rawCourse);


    const startIndex =
      match.index ?? 0;


    const endIndex =
      i + 1 <
      courseMatches.length
        ? courseMatches[i + 1].index ??
          tableText.length
        : tableText.length;


    const courseBlock =
      tableText.substring(
        startIndex,
        endIndex
      );


    /*
     * Determine whether this is a lab.
     */

    /*
     * Get the base course code.
     *
     * CSE302 Lab → CSE302
     * CSE325     → CSE325
     */
    const baseCourse =
      course
        .replace(/\s+Lab$/i, "")
        .toUpperCase();


    const section = sectionMap.get(baseCourse);


    /*
     * Extract all schedules belonging
     * to this course.
     */
    const entries =
      expandScheduleText(
        courseBlock,
        course,
        section
      );


    results.push(...entries);
  }


  return removeDuplicates(
    results
  );
}


/* ============================================================
   PDF TEXT EXTRACTION
============================================================ */

async function extractPdfText(
  file: File
): Promise<string> {
  const arrayBuffer =
    await file.arrayBuffer();


  const pdf =
    await pdfjsLib.getDocument({
      data: arrayBuffer,
    }).promise;


  let fullText = "";


  for (
    let pageNumber = 1;
    pageNumber <= pdf.numPages;
    pageNumber++
  ) {
    const page =
      await pdf.getPage(
        pageNumber
      );


    const content =
      await page.getTextContent();


    const pageText =
      content.items
        .map((item) =>
          "str" in item
            ? item.str
            : ""
        )
        .join(" ");


    fullText +=
      ` ${pageText}`;
  }


  return fullText;
}


/* ============================================================
   EXCEL PARSER
============================================================ */

function parseExcelWorkbook(
  workbook: XLSX.WorkBook
): ClassEntry[] {
  const results: ClassEntry[] = [];

  /*
   * ============================================================
   * FIRST PASS
   *
   * Build:
   *
   * CSE246  → 2
   * CSE302  → 10
   * MKT7101 → 9
   *
   * This lets both normal courses and labs use
   * the same section.
   * ============================================================
   */

  const sectionMap =
    new Map<string, string>();

  for (
    const sheetName
    of workbook.SheetNames
  ) {
    const sheet =
      workbook.Sheets[
        sheetName
      ];

    const rows =
      XLSX.utils.sheet_to_json(
        sheet,
        {
          header: 1,
          defval: "",
        }
      ) as unknown[][];

    for (
      const row of rows
    ) {
      const cells =
        row
          .map((cell) =>
            String(cell).trim()
          )
          .filter(Boolean);

      if (
        cells.length === 0
      ) {
        continue;
      }

      const rowText =
        cells.join(" ");

      /*
       * Find course code.
       */
      const courseMatch =
        rowText.match(
          /\b([A-Z]{2,4}\d{3,4})\b/i
        );

      if (!courseMatch) {
        continue;
      }

      const baseCourse =
        courseMatch[1].toUpperCase();

      /*
       * Find the position of the course
       * inside the row.
       */
      const courseIndex =
        cells.findIndex(
          (cell) =>
            cell
              .toUpperCase()
              .includes(baseCourse)
        );

      /*
       * Normally Excel looks like:
       *
       * CSE246 | 2 | 4.00 | ...
       */
      if (
        courseIndex !== -1
      ) {
        const possibleSection =
          cells[
            courseIndex + 1
          ];

        if (
          possibleSection &&
          /^\d+$/.test(
            possibleSection
          )
        ) {
          sectionMap.set(
            baseCourse,
            possibleSection
          );

          continue;
        }
      }

      /*
       * Fallback if Excel flattened the row.
       *
       * Example:
       *
       * CSE246 2 4.00 ...
       */
      const flexibleMatch =
        rowText.match(
          new RegExp(
            `\\b${baseCourse}\\b\\s+(\\d+)\\s+\\d+(?:\\.\\d+)?`,
            "i"
          )
        );

      if (
        flexibleMatch
      ) {
        sectionMap.set(
          baseCourse,
          flexibleMatch[1]
        );
      }
    }
  }


  /*
   * ============================================================
   * SECOND PASS
   *
   * Parse the actual schedule.
   * ============================================================
   */

  for (
    const sheetName
    of workbook.SheetNames
  ) {
    const sheet =
      workbook.Sheets[
        sheetName
      ];

    const rows =
      XLSX.utils.sheet_to_json(
        sheet,
        {
          header: 1,
          defval: "",
        }
      ) as unknown[][];

    let currentCourse = "";
    let currentSection:
      string | undefined;


    for (
      const row of rows
    ) {
      const cells =
        row
          .map((cell) =>
            String(cell).trim()
          )
          .filter(Boolean);

      if (
        cells.length === 0
      ) {
        continue;
      }

      const rowText =
        cells.join(" ");


      /*
       * Find course code.
       */
      const courseMatch =
        rowText.match(
          /\b([A-Z]{2,4}\d{3,4})\b(?:\s+Lab)?/i
        );


      if (courseMatch) {
        const baseCourse =
          courseMatch[1]
            .toUpperCase();

        /*
         * Determine whether this is a lab.
         */
        const hasLab =
          /\bLab\b/i.test(
            rowText
          );

        currentCourse =
          hasLab
            ? `${baseCourse} Lab`
            : baseCourse;

        /*
         * IMPORTANT:
         *
         * Both the lecture and lab get the
         * section from the same course map.
         *
         * CSE246     → Section 2
         * CSE246 Lab → Section 2
         */
        currentSection =
          sectionMap.get(
            baseCourse
          );
      }


      /*
       * Ignore rows before the first course.
       */
      if (
        !currentCourse
      ) {
        continue;
      }


      /*
       * Find schedule.
       */
      const rowScheduleRegex =
        /([ASMTWRF]+)\s+(\d{1,2}:\d{2}\s*(?:AM|PM))-(\d{1,2}:\d{2}\s*(?:AM|PM))/gi;

      rowScheduleRegex.lastIndex =
        0;

      const scheduleMatches = [
        ...rowText.matchAll(
          rowScheduleRegex
        ),
      ];


      for (
        const match
        of scheduleMatches
      ) {
        const dayCodes =
          match[1].toUpperCase();

        const start =
          normalizeTime(
            match[2]
          );

        const end =
          normalizeTime(
            match[3]
          );


        let room = "";


        /*
         * Find the cell containing
         * the schedule.
         */
        rowScheduleRegex.lastIndex =
          0;

        const timeCellIndex =
          cells.findIndex(
            (cell) => {
              rowScheduleRegex.lastIndex =
                0;

              return rowScheduleRegex.test(
                cell
              );
            }
          );


        /*
         * Usually the room is the
         * cell immediately after time.
         */
        if (
          timeCellIndex !== -1 &&
          cells[
            timeCellIndex + 1
          ]
        ) {
          room =
            cells[
              timeCellIndex + 1
            ];
        }


        /*
         * Fallback room extraction.
         */
        if (!room) {
          const afterSchedule =
            rowText
              .substring(
                (match.index ?? 0) +
                  match[0].length
              )
              .trim();

          const roomMatch =
            afterSchedule.match(
              /^([A-Za-z0-9-]+(?:\s+\([^)]*\))?)/
            );

          room =
            roomMatch?.[1] ??
            "";
        }


        /*
         * Expand combined days.
         *
         * MW → Monday + Wednesday
         * S  → Sunday
         * R  → Thursday
         */
        for (
          const code
          of dayCodes
        ) {
          if (
            !dayNames[code]
          ) {
            continue;
          }

          results.push({
            course:
              currentCourse,

            section:
              currentSection,

            day:
              dayNames[code],

            start,

            end,

            room,
          });
        }
      }
    }
  }


  return removeDuplicates(
    results
  );
}

/* ============================================================
   FILE PARSER
============================================================ */

async function parseFile(
  file: File,
  onStudentName?: (name: string) => void,
  onStudentFullName?: (name: string) => void,
  onStudentId?: (id: string) => void
): Promise<ClassEntry[]> {
  const extension =
    file.name
      .split(".")
      .pop()
      ?.toLowerCase();


  if (
    extension === "pdf"
  ) {
    const text =
      await extractPdfText(
        file
      );


    return parsePdfText(
      text
    );
  }


if (
  extension === "xlsx" ||
  extension === "xls"
) {
  const buffer =
    await file.arrayBuffer();

  const workbook =
    XLSX.read(
      buffer,
      {
        type: "array",
      }
    );

  // Get text from the first Excel sheet
const studentInfo =
  extractStudentInfoFromExcel(
    workbook
  );

onStudentName?.(
  studentInfo.firstName
);

onStudentFullName?.(
  studentInfo.fullName
);

onStudentId?.(
  studentInfo.studentId
);

  // Existing Excel schedule parsing
  return parseExcelWorkbook(
    workbook
  );
}


  throw new Error(
    "Unsupported file type."
  );
}


/* ============================================================
   APP
============================================================ */

function App() {
    const [now, setNow] =
    useState(new Date());

  useEffect(() => {
    const timer =
      window.setInterval(() => {
        setNow(new Date());
      }, 30000);

    return () =>
      window.clearInterval(timer);
  }, []);


  const [fileName, setFileName] = useState(() => {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);

    if (!saved) {
      return "";
    }

    const data = JSON.parse(saved);

    return data.fileName || "";
  } catch {
    return "";
  }
});

const [studentName, setStudentName] = useState(() => {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);

    if (!saved) {
      return "Student";
    }

    const data = JSON.parse(saved);

    return data.studentName || "Student";
  } catch {
    return "Student";
  }
});

const [studentFullName, setStudentFullName] =
  useState(() => {
    try {
      const saved =
        localStorage.getItem(STORAGE_KEY);

      if (!saved) {
        return "Student";
      }

      const data = JSON.parse(saved);

      return data.studentFullName || "Student";
    } catch {
      return "Student";
    }
  });

const [studentId, setStudentId] =
  useState(() => {
    try {
      const saved =
        localStorage.getItem(STORAGE_KEY);

      if (!saved) {
        return "";
      }

      const data = JSON.parse(saved);

      return data.studentId || "";
    } catch {
      return "";
    }
  });

const [classes, setClasses] = useState<ClassEntry[]>(() => {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);

    if (!saved) {
      return [];
    }

    const data = JSON.parse(saved);

    return Array.isArray(data.classes)
      ? data.classes
      : [];
  } catch {
    return [];
  }
});

useEffect(() => {
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({
      fileName,
      studentName,
      studentFullName,
      studentId,
      classes,
    })
  );
}, [
  fileName,
  studentName,
  studentFullName,
  studentId,
  classes,
]);

  const [loading, setLoading] =
    useState(false);

  const [error, setError] =
    useState("");

  const [editingIndex, setEditingIndex] =
    useState<number | null>(null);

  const [addingClass, setAddingClass] =
  useState(false);

  const [showExportModal, setShowExportModal] =
  useState(false);

const [newClass, setNewClass] =
  useState<ClassEntry>({
    course: "",
    section: "",
    day: "Saturday",
    start: "",
    end: "",
    room: "",
  });  
    
    const [editData, setEditData] =
  useState<ClassEntry | null>(null);

    const today =
    jsDayToName[now.getDay()];

  const todayClasses =
    classes
      .filter(
        (item) =>
          item.day === today
      )
      .sort(
        (a, b) =>
          timeToMinutes(a.start) -
          timeToMinutes(b.start)
      );

  const currentMinutes =
    now.getHours() * 60 +
    now.getMinutes();


let runningClass: ClassEntry | null = null;

const upcomingTodayClasses: ClassEntry[] = [];

for (const item of todayClasses) {
  const start = timeToMinutes(item.start);
  const end = timeToMinutes(item.end);

  if (
    currentMinutes >= start &&
    currentMinutes < end
  ) {
    runningClass = item;
  }

  if (currentMinutes < end) {
    upcomingTodayClasses.push(item);
  }
}

const nextUpcomingIndex =
  upcomingTodayClasses.findIndex(
    (item) =>
      timeToMinutes(item.start) >
      currentMinutes
  );

let nextDayClass: ClassEntry | null = null;
let nextDayName: string | null = null;

const todayIndex = days.indexOf(today);

for (let offset = 1; offset <= 7; offset++) {
  const day =
    days[(todayIndex + offset) % days.length];

  const futureClasses =
    classes
      .filter((item) => item.day === day)
      .sort(
        (a, b) =>
          timeToMinutes(a.start) -
          timeToMinutes(b.start)
      );

  if (futureClasses.length > 0) {
    nextDayClass = futureClasses[0];
    nextDayName = day;
    break;
  }
}


  const handleFile = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file =
      event.target.files?.[0];


    if (!file) {
      return;
    }

    const extension =
  file.name
    .split(".")
    .pop()
    ?.toLowerCase();


    setFileName(
      file.name
    );

    setClasses([]);

    setError("");

    setLoading(true);


try {
  let parsedClasses: ClassEntry[];

  if (extension === "pdf") {
    const pdfText =
      await extractPdfText(file);

    const extractedName =
      extractStudentName(pdfText);

    const extractedFullName =
      extractFullStudentName(pdfText);

    const extractedId =
      extractStudentId(pdfText);

    console.log(
      "===== STUDENT INFO ====="
    );

    console.log(
      "First name:",
      extractedName
    );

    console.log(
      "Full name:",
      extractedFullName
    );

    console.log(
      "Student ID:",
      extractedId
    );

    console.log(
      "========================"
    );

    setStudentName(
      extractedName
    );

    setStudentFullName(
      extractedFullName
    );

    setStudentId(
      extractedId
    );

    parsedClasses =
      parsePdfText(pdfText);

} else {
  parsedClasses =
    await parseFile(
      file,
      setStudentName,
      setStudentFullName,
      setStudentId
    );
}

  if (
    parsedClasses.length ===
    0
  ) {
    setError(
      "No class schedule could be found in this file."
    );
  } else {
    setClasses(
      parsedClasses
    );
  }

} catch (err) {
  console.error(err);

  setError(
    "Something went wrong while reading the file."
  );
} finally {
  setLoading(false);
}
  };

const exportRoutinePDF = async () => {
  if (classes.length === 0) {
    alert("There is no routine to export.");
    return;
  }

  try {
    /* =========================
       LOAD LOGO
    ========================= */

    const logoUrl = new URL(
      `${import.meta.env.BASE_URL}logo.svg`,
      window.location.origin
    ).href;

    const logoResponse = await fetch(logoUrl);

    if (!logoResponse.ok) {
      throw new Error(
        `Could not load logo.svg: ${logoResponse.status}`
      );
    }

const logoSvgText =
  await logoResponse.text();

/* =========================
   CONVERT SVG → PNG
========================= */

const logoBlob =
  new Blob(
    [logoSvgText],
    {
      type: "image/svg+xml",
    }
  );

const logoObjectUrl =
  URL.createObjectURL(logoBlob);

const logoImage =
  new Image();

logoImage.src =
  logoObjectUrl;

await new Promise<void>(
  (resolve, reject) => {
    logoImage.onload = () =>
      resolve();

    logoImage.onerror = () =>
      reject(
        new Error(
          "Could not render logo.svg"
        )
      );
  }
);

const logoCanvas =
  document.createElement("canvas");

const logoWidth = 300;

const logoHeight =
  logoImage.naturalHeight *
  (logoWidth /
    logoImage.naturalWidth);

logoCanvas.width =
  logoWidth;

logoCanvas.height =
  logoHeight;

const logoContext =
  logoCanvas.getContext("2d");

if (!logoContext) {
  throw new Error(
    "Could not create logo canvas"
  );
}

logoContext.drawImage(
  logoImage,
  0,
  0,
  logoWidth,
  logoHeight
);

const logoDataUrl =
  logoCanvas.toDataURL(
    "image/png"
  );

URL.revokeObjectURL(
  logoObjectUrl
);
    /* =========================
       CREATE PDF
    ========================= */

    const doc = new jsPDF({
      orientation: "landscape",
      unit: "mm",
      format: "a4",
    });

    const pageWidth = 297;
    const pageHeight = 210;

    const margin = 10;
    const contentWidth =
      pageWidth - margin * 2;

    const columnGap = 2;

    const columnWidth =
      (contentWidth -
        columnGap * 6) /
      7;

    /* =========================
       HEADER
    ========================= */

    // Logo
const pdfLogoWidth = 22;

const pdfLogoHeight =
  pdfLogoWidth *
  (logoImage.naturalHeight /
    logoImage.naturalWidth);

doc.addImage(
  logoDataUrl,
  "PNG",
  margin,
  7,
  pdfLogoWidth,
  pdfLogoHeight
);

    // Title
    doc.setFont(
      "helvetica",
      "bold"
    );

    doc.setFontSize(22);

    doc.text(
      "EWOrganizer",
      margin + 27,
      17
    );

    // Subtitle
    doc.setFont(
      "helvetica",
      "normal"
    );

    doc.setFontSize(10);

    doc.text(
      "Weekly Routine",
      margin + 27,
      23
    );

    /* =========================
       STUDENT INFO
    ========================= */

    doc.setFontSize(10);

    doc.text(
      `Student: ${studentFullName}`,
      pageWidth - margin,
      17,
      {
        align: "right",
      }
    );

    if (studentId) {
      doc.setFontSize(9);

      doc.text(
        `ID: ${studentId}`,
        pageWidth - margin,
        23,
        {
          align: "right",
        }
      );
    }

    /* =========================
       ROUTINE GRID
    ========================= */

    const gridTop = 31;
    const headerHeight = 10;

    days.forEach(
      (day, dayIndex) => {
        const x =
          margin +
          dayIndex *
            (columnWidth + columnGap);

        /* =========================
           DAY HEADER
        ========================= */

        doc.setFillColor(
          0,
          0,
          0
        );

        doc.roundedRect(
          x,
          gridTop,
          columnWidth,
          headerHeight,
          2,
          2,
          "F"
        );

        doc.setTextColor(
          255,
          255,
          255
        );

        doc.setFont(
          "helvetica",
          "bold"
        );

        doc.setFontSize(8);

        doc.text(
          day,
          x +
            columnWidth / 2,
          gridTop + 6.5,
          {
            align: "center",
          }
        );

        doc.setTextColor(
          0,
          0,
          0
        );

        /* =========================
           CLASSES
        ========================= */

        const dayClasses =
          classes
            .filter(
              (item) =>
                item.day === day
            )
            .sort(
              (a, b) =>
                a.start.localeCompare(
                  b.start
                )
            );

        let cardY =
          gridTop +
          headerHeight +
          3;

        dayClasses.forEach(
          (item) => {
            /* =========================
               ROOM TEXT
            ========================= */

            const roomText =
              `Room: ${item.room}`;

            doc.setFont(
              "helvetica",
              "normal"
            );

            doc.setFontSize(6.5);

            const roomLines =
              doc.splitTextToSize(
                roomText,
                columnWidth - 5
              );

            /* =========================
               CARD HEIGHT
            ========================= */

            const cardHeight =
              (item.section
                ? 30
                : 26) +
              Math.max(
                0,
                roomLines.length - 1
              ) *
                4;

            /* =========================
               CARD
            ========================= */

            doc.setDrawColor(
              225,
              225,
              228
            );

            doc.setFillColor(
              255,
              255,
              255
            );

            doc.roundedRect(
              x,
              cardY,
              columnWidth,
              cardHeight,
              2,
              2,
              "FD"
            );

            /* =========================
               COURSE
            ========================= */

            doc.setFont(
              "helvetica",
              "bold"
            );

            doc.setFontSize(7.5);

            doc.text(
              item.course,
              x + 2.5,
              cardY + 6
            );

            /* =========================
               SECTION
            ========================= */

            if (item.section) {
              doc.setFont(
                "helvetica",
                "normal"
              );

              doc.setFontSize(6.5);

              doc.text(
                `Section ${item.section}`,
                x + 2.5,
                cardY + 11
              );
            }

            /* =========================
               TIME
            ========================= */

            const timeY =
              item.section
                ? cardY + 17
                : cardY + 12;

            doc.setFont(
              "helvetica",
              "normal"
            );

            doc.setFontSize(6.5);

            doc.text(
              `${item.start} - ${item.end}`,
              x + 2.5,
              timeY
            );

            /* =========================
               ROOM
            ========================= */

            doc.text(
              roomLines,
              x + 2.5,
              timeY + 5
            );

            cardY +=
              cardHeight + 3;
          }
        );
      }
    );

    /* =========================
       FOOTER
    ========================= */

    doc.setFont(
      "helvetica",
      "normal"
    );

    doc.setFontSize(7);

    doc.text(
      `Generated by EWOrganizer • ${classes.length} class sessions`,
      pageWidth / 2,
      pageHeight - 7,
      {
        align: "center",
      }
    );

    /* =========================
       DOWNLOAD
    ========================= */

    doc.save(
      "EWOrganizer-Routine.pdf"
    );

  } catch (error) {
    console.error(
      "PDF export failed:",
      error
    );

    alert(
      "Something went wrong while exporting the PDF."
    );
  }
};

const exportRoutinePNG = async () => {
  if (classes.length === 0) {
    alert("There is no routine to export.");
    return;
  }

  const routineElement =
    document.querySelector(".routine-grid") as HTMLElement | null;

  if (!routineElement) {
    alert("Could not find the routine to export.");
    return;
  }

  try {
    /* =========================
       LOAD EXISTING LOGO SVG
    ========================= */

    const logoUrl = new URL(
      `${import.meta.env.BASE_URL}logo.svg`,
      window.location.origin
    ).href;

    const logoResponse = await fetch(logoUrl);

    if (!logoResponse.ok) {
      throw new Error(
        `Could not load logo.svg: ${logoResponse.status}`
      );
    }

    const logoSvgText =
      await logoResponse.text();

const logoBlob =
  new Blob(
    [logoSvgText],
    { type: "image/svg+xml" }
  );

const logoObjectUrl =
  URL.createObjectURL(logoBlob);

const logoImage =
  new Image();

logoImage.src =
  logoObjectUrl;

await new Promise<void>((resolve, reject) => {
  logoImage.onload = () =>
    resolve();

  logoImage.onerror = () =>
    reject(
      new Error("Could not render logo.svg")
    );
});

const logoCanvas =
  document.createElement("canvas");

logoCanvas.width = 300;
logoCanvas.height = 300;

const logoContext =
  logoCanvas.getContext("2d");

if (!logoContext) {
  throw new Error(
    "Could not create logo canvas"
  );
}

logoContext.drawImage(
  logoImage,
  0,
  0,
  300,
  300
);

URL.revokeObjectURL(
  logoObjectUrl
); 

    /* =========================
       CREATE ROUTINE IMAGE
    ========================= */

    const canvas =
      await html2canvas(routineElement, {
        backgroundColor: "#f6f7f9",
        scale: 2,
        useCORS: true,

        onclone: (clonedDocument) => {
          const clonedGrid =
            clonedDocument.querySelector(
              ".routine-grid"
            ) as HTMLElement | null;

          if (!clonedGrid) return;

          /* =========================
             EXPORT SIZE
          ========================= */

          clonedGrid.style.width = "1400px";
          clonedGrid.style.maxWidth = "1400px";
          clonedGrid.style.minWidth = "1400px";

          clonedGrid.style.display = "grid";

          clonedGrid.style.gridTemplateColumns =
            "repeat(7, minmax(0, 1fr))";

          clonedGrid.style.gap = "10px";

          clonedGrid.style.paddingTop = "105px";
          clonedGrid.style.paddingBottom = "55px";
          clonedGrid.style.paddingLeft = "0";
          clonedGrid.style.paddingRight = "0";

          clonedGrid.style.boxSizing =
            "border-box";

          clonedGrid.style.position =
            "relative";

          /* =========================
             HIDE WEB-ONLY BUTTONS
          ========================= */

          clonedGrid
            .querySelectorAll(
              ".edit-btn, .delete-btn"
            )
            .forEach((button) => {
              (
                button as HTMLElement
              ).style.display = "none";
            });

          /* =========================
             EXPORT HEADER
          ========================= */

          const header =
            clonedDocument.createElement("div");

          header.style.position =
            "absolute";

          header.style.top = "12px";
          header.style.left = "24px";
          header.style.right = "24px";

          header.style.display = "flex";

          header.style.alignItems =
            "flex-start";

          header.style.justifyContent =
            "space-between";

          /* =========================
             LEFT BRANDING
          ========================= */

          const left =
            clonedDocument.createElement("div");

          left.style.display = "flex";
          left.style.flexDirection =
            "column";

          left.style.alignItems =
            "flex-start";

          /* =========================
   LOGO + TEXT
========================= */

const branding =
  clonedDocument.createElement("div");

branding.style.display = "flex";
branding.style.alignItems = "center";
branding.style.gap = "8px";

/* =========================
   INLINE SVG LOGO
========================= */

const parser =
  new DOMParser();

const parsedLogo =
  parser.parseFromString(
    logoSvgText,
    "image/svg+xml"
  );

const logoSvg =
  parsedLogo.documentElement;

logoSvg.setAttribute(
  "width",
  "88"
);

logoSvg.setAttribute(
  "height",
  "88"
);

logoSvg.style.width = "88px";
logoSvg.style.height = "88px";
logoSvg.style.display = "block";
logoSvg.style.flexShrink = "0";

/* Import SVG into cloned document */

const importedLogo =
  clonedDocument.importNode(
    logoSvg,
    true
);

/* =========================
   TITLE + SUBTITLE COLUMN
========================= */

const textBlock =
  clonedDocument.createElement("div");

textBlock.style.display = "flex";
textBlock.style.flexDirection =
  "column";

textBlock.style.alignItems =
  "flex-start";

const title =
  clonedDocument.createElement("div");

title.textContent =
  "EWOrganizer";

title.style.fontSize = "30px";
title.style.fontWeight = "700";
title.style.lineHeight = "1";
title.style.color = "#18181b";

textBlock.appendChild(
  title
);

/* =========================
   SUBTITLE
========================= */

const subtitle =
  clonedDocument.createElement("div");

subtitle.textContent =
  "Weekly Routine";

subtitle.style.marginTop = "7px";
subtitle.style.fontSize = "14px";
subtitle.style.fontWeight = "500";
subtitle.style.lineHeight = "1.2";
subtitle.style.color = "#71717a";

textBlock.appendChild(
  subtitle
);

/* =========================
   COMBINE LOGO + TEXT
========================= */

branding.appendChild(
  importedLogo
);

branding.appendChild(
  textBlock
);

left.appendChild(
  branding
);

          /* =========================
             RIGHT STUDENT INFO
          ========================= */

          const right =
            clonedDocument.createElement(
              "div"
            );

          right.style.display = "flex";

          right.style.flexDirection =
            "column";

          right.style.alignItems =
            "flex-end";

          right.style.textAlign =
            "right";

          const student =
            clonedDocument.createElement(
              "div"
            );

          student.textContent =
            `Student: ${studentFullName}`;

          student.style.fontSize = "19px";

          student.style.fontWeight =
            "700";

          student.style.lineHeight =
            "1.2";

          student.style.color =
            "#18181b";

          right.appendChild(
            student
          );

          if (studentId) {
            const id =
              clonedDocument.createElement(
                "div"
              );

            id.textContent =
              `ID: ${studentId}`;

            id.style.marginTop = "8px";

            id.style.fontSize = "16px";

            id.style.fontWeight =
              "500";

            id.style.lineHeight =
              "1.2";

            id.style.color =
              "#71717a";

            right.appendChild(
              id
            );
          }

          header.appendChild(
            left
          );

          header.appendChild(
            right
          );

          clonedGrid.appendChild(
            header
          );

          /* =========================
             FOOTER
          ========================= */

          const footer =
            clonedDocument.createElement(
              "div"
            );

          footer.textContent =
            `Generated by EWOrganizer • ${classes.length} class sessions`;

          footer.style.position =
            "absolute";

          footer.style.bottom = "18px";

          footer.style.left = "0";
          footer.style.right = "0";

          footer.style.textAlign =
            "center";

          footer.style.fontSize = "11px";

          footer.style.fontWeight =
            "500";

          footer.style.color =
            "#a1a1aa";

          clonedGrid.appendChild(
            footer
          );
        },
      });

    /* =========================
       DOWNLOAD
    ========================= */

    const link =
      document.createElement("a");

    link.download =
      "EWOrganizer-Routine.png";

    link.href =
      canvas.toDataURL("image/png");

    link.click();

  } catch (error) {
    console.error(
      "PNG export failed:",
      error
    );

    alert(
      "Something went wrong while exporting the routine."
    );
  }
};

  return (
<div className="app">

  <header className="header">

    <div className="brand-row">

      <img
        src={`${import.meta.env.BASE_URL}logo.svg`}
        alt="ClassMate logo"
        className="brand-logo"
      />

      <div>
<h1>
  <span className="ewu-text">EW</span>Organizer
</h1>

        <p>
          Organize your semester schedule
        </p>
      </div>

    </div>


<div className="semester">
  {studentName}
</div>

      </header>


      <main>

        <section className="upload-section">

          <div className="upload-icon">
            📚
          </div>


          <h2>
            Import your semester slip
          </h2>


          <p>
            Upload your EWU advising slip
            in PDF or Excel format.
          </p>


          <label className="upload-button">

            Choose PDF or Excel

            <input
              type="file"
              accept=".pdf,.xlsx,.xls"
              onChange={
                handleFile
              }
              hidden
            />

          </label>


          {fileName && (
            <div className="selected-file">
              ✓ {fileName}
            </div>
          )}


          {loading && (
            <div className="status">
              Reading your routine...
            </div>
          )}


          {error && (
            <div className="error">
              ⚠ {error}
            </div>
          )}

        </section>


{classes.length > 0 && (
  <>
    {/* Today / Next Class */}
<section className="today-panel">

  {/* TODAY SUMMARY */}
  <div className="today-info">
    <div className="today-date">
      {formatToday(now)}
    </div>

    <div className="today-count">
      {todayClasses.length === 0
        ? "No classes today"
        : `${todayClasses.length} ${
            todayClasses.length === 1
              ? "class"
              : "classes"
          } today`}
    </div>
  </div>


{/* TODAY'S CLASSES */}

<div className="today-classes">

  {upcomingTodayClasses.map((item, index) => {

    const isRunning = runningClass === item;

    return (
      <div
className={`timeline-class ${
  isRunning
    ? "running-class"
    : index === nextUpcomingIndex
    ? "next-class"
    : "upcoming-class"
}`}
        key={`${item.course}-${item.start}-${index}`}
      >

<div className="next-label">
  {isRunning
    ? "RUNNING NOW"
    : index === nextUpcomingIndex
    ? "NEXT CLASS"
    : "UPCOMING"}
</div>

        <div className="next-course">
          {item.course}
          {item.section &&
            ` · Section ${item.section}`}
        </div>

        <div className="next-details">
          {item.start}
          {" – "}
          {item.end}
          {" · "}
          {item.room}
        </div>

      </div>
    );
  })}

  {todayClasses.length > 0 &&
    upcomingTodayClasses.length === 0 && (
      <div className="today-finished">
        <div className="next-label">
          TODAY
        </div>

        <div className="next-course">
          No more classes today
        </div>
      </div>
    )}

  {/* NEXT DAY */}
  {nextDayClass && (
    <div className="next-day-class">

      <div className="next-label">
        NEXT · {nextDayName?.toUpperCase()}
      </div>

      <div className="next-course">
        {nextDayClass.course}
        {nextDayClass.section &&
          ` · Section ${nextDayClass.section}`}
      </div>

      <div className="next-details">
        {nextDayClass.start}
        {" – "}
        {nextDayClass.end}
        {" · "}
        {nextDayClass.room}
      </div>

    </div>
  )}

</div>


</section>

{/* Weekly Routine */}
<section className="routine-section">

<div className="section-header">

  <div>

    <h2>
      Weekly Routine
    </h2>

    <p>
      {classes.length} class sessions
      found
    </p>

  </div>

<div className="routine-actions">
  <button
    className="export-btn"
  onClick={() => {
    setShowExportModal(true);
  }}
>
  Export
</button>

    <button
      className="add-class-btn"
      onClick={() => {
        setAddingClass(true);
      }}
    >
      + Add Class
    </button>
  </div>

</div>

{showExportModal && (
  <div
    className="modal-overlay"
    onClick={() => setShowExportModal(false)}
  >
    <div
      className="add-class-modal export-modal"
      onClick={(e) => e.stopPropagation()}
    >
      <div className="modal-header">
        <div>
          <h3>Export Routine</h3>
          <p>Choose what you want to do</p>
        </div>

        <button
          className="modal-close"
          onClick={() =>
            setShowExportModal(false)
          }
        >
          ✕
        </button>
      </div>

      <div className="export-options">

        <button
          className="export-option"
onClick={async () => {
  setShowExportModal(false);
  await exportRoutinePDF();
}}
        >
          <span className="export-option-icon">
            📄
          </span>

          <span>
            <strong>Export as PDF</strong>
            <small>
              Download your routine as a PDF
            </small>
          </span>
        </button>

        <button
          className="export-option"
          onClick={() => {
            setShowExportModal(false);
            exportRoutinePNG();
          }}
        >
          <span className="export-option-icon">
            🖼️
          </span>

          <span>
            <strong>Export as PNG</strong>
            <small>
              Save your routine as an image
            </small>
          </span>
        </button>

        <button
          className="export-option"
          onClick={() => {
            setShowExportModal(false);
            alert("Share feature coming next.");
          }}
        >
          <span className="export-option-icon">
            🔗
          </span>

          <span>
            <strong>Share</strong>
            <small>
              Share your routine with others
            </small>
          </span>
        </button>

      </div>

      <div className="modal-actions">
        <button
          className="modal-cancel"
          onClick={() =>
            setShowExportModal(false)
          }
        >
          Cancel
        </button>
      </div>

    </div>
  </div>
)}

{addingClass && (
  <div
    className="modal-overlay"
    onClick={() => setAddingClass(false)}
  >
    <div
      className="add-class-modal"
      onClick={(e) => e.stopPropagation()}
    >

      <div className="modal-header">
        <div>
          <h3>Add New Class</h3>
          <p>Enter the class details below</p>
        </div>

        <button
          className="modal-close"
          onClick={() => setAddingClass(false)}
        >
          ✕
        </button>
      </div>

      <div className="modal-form">

        <input
          placeholder="Course (e.g. CSE302)"
          value={newClass.course}
          onChange={(e) =>
            setNewClass({
              ...newClass,
              course: e.target.value,
            })
          }
        />

        <input
          placeholder="Section"
          value={newClass.section}
          onChange={(e) =>
            setNewClass({
              ...newClass,
              section: e.target.value,
            })
          }
        />

<select
  value={newClass.day}
  onChange={(e) =>
    setNewClass({
      ...newClass,
      day: e.target.value,
    })
  }
>
  <option value="Saturday">Saturday</option>
  <option value="Sunday">Sunday</option>
  <option value="Monday">Monday</option>
  <option value="Tuesday">Tuesday</option>
  <option value="Wednesday">Wednesday</option>
  <option value="Thursday">Thursday</option>
  <option value="Friday">Friday</option>
</select>

        <input
          placeholder="Start (e.g. 10:00AM)"
          value={newClass.start}
          onChange={(e) =>
            setNewClass({
              ...newClass,
              start: e.target.value,
            })
          }
        />

        <input
          placeholder="End (e.g. 11:30AM)"
          value={newClass.end}
          onChange={(e) =>
            setNewClass({
              ...newClass,
              end: e.target.value,
            })
          }
        />

        <input
          placeholder="Room"
          value={newClass.room}
          onChange={(e) =>
            setNewClass({
              ...newClass,
              room: e.target.value,
            })
          }
        />

      </div>

      <div className="modal-actions">

        <button
          className="modal-cancel"
          onClick={() => setAddingClass(false)}
        >
          Cancel
        </button>

        <button
          className="modal-add"
          onClick={() => {
            setClasses([
              ...classes,
              newClass,
            ]);

            setNewClass({
              course: "",
              section: "",
              day: "Saturday",
              start: "",
              end: "",
              room: "",
            });

            setAddingClass(false);
          }}
        >
          ＋ Add Class
        </button>

      </div>

    </div>
  </div>
)}


<div className="routine-grid">

        {days.map(
          (day) => {

            const dayClasses =
              classes
                .filter(
                  (item) =>
                    item.day === day
                )
                .sort(
                  (a, b) =>
                    a.start.localeCompare(
                      b.start
                    )
                );


            return (

              <div
                className="day-column"
                key={day}
              >

<div
  className={`day-header ${
    day === today
      ? "today-header"
      : ""
  }`}
>
  {day}
</div>


                {dayClasses.length ===
                0 ? (

                  <div className="no-class">
                    No classes
                  </div>

                ) : (

                  dayClasses.map(
                    (
                      item,
                      index
                    ) => (

                      <div
                        className="class-card"
                        key={`${item.course}-${item.day}-${item.start}-${item.room}-${index}`}
                      >

                        {editingIndex === classes.indexOf(item) ? (
<div className="edit-form">

<select
  value={editData?.day || "Saturday"}
  onChange={(e) =>
    setEditData({
      ...editData!,
      day: e.target.value,
    })
  }
>
  <option value="Saturday">Saturday</option>
  <option value="Sunday">Sunday</option>
  <option value="Monday">Monday</option>
  <option value="Tuesday">Tuesday</option>
  <option value="Wednesday">Wednesday</option>
  <option value="Thursday">Thursday</option>
  <option value="Friday">Friday</option>
</select>
  
  <input
    value={editData?.course || ""}
    onChange={(e) =>
      setEditData({
        ...editData!,
        course: e.target.value,
      })
    }
  />

  <input
    value={editData?.section || ""}
    placeholder="Section"
    onChange={(e) =>
      setEditData({
        ...editData!,
        section: e.target.value,
      })
    }
  />

  <input
    value={editData?.start || ""}
    onChange={(e) =>
      setEditData({
        ...editData!,
        start: e.target.value,
      })
    }
  />

  <input
    value={editData?.end || ""}
    onChange={(e) =>
      setEditData({
        ...editData!,
        end: e.target.value,
      })
    }
  />

  <input
    value={editData?.room || ""}
    placeholder="Room"
    onChange={(e) =>
      setEditData({
        ...editData!,
        room: e.target.value,
      })
    }
  />

  <button
    className="save-edit-btn"
    onClick={() => {
      if (
        editingIndex !== null &&
        editData
      ) {
        const updated = [...classes];

        updated[editingIndex] = editData;

        setClasses(updated);
      }

      setEditingIndex(null);
      setEditData(null);
    }}
  >
    ✓ Done
  </button>

</div>
) : (
  <>
    <div className="course-name">
      {item.course}
    </div>

    {item.section && (
      <div className="section">
        Section {item.section}
      </div>
    )}

    <div className="class-time">
      {item.start} – {item.end}
    </div>

    <div className="class-room">
      📍 {item.room}
    </div>

<button
  className="edit-btn"
  onClick={() => {
    setEditingIndex(
      classes.indexOf(item)
    );
    setEditData({ ...item });
  }}
>
  Edit
</button>

<button
  className="delete-btn"
  onClick={() => {
    const index = classes.indexOf(item);

    if (
      window.confirm(
        `Delete ${item.course}${item.section ? ` Section ${item.section}` : ""}?`
      )
    ) {
      const updated = [...classes];
      updated.splice(index, 1);
      setClasses(updated);
    }
  }}
>
  🗑 Delete
</button>
  </>
)}

                      </div>

                    )
                  )

                )}

              </div>

            );

          }
        )}

      </div>

    </section>
  </>
)}

      </main>


      <footer>
        EWOrganizer • by Double M
      </footer>

    </div>
  );
}



export default App;