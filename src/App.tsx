import { useEffect, useState } from "react";
import * as pdfjsLib from "pdfjs-dist";
import * as XLSX from "xlsx";
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

function getFirstName(
  fullName: string
): string {
  const nameParts =
    fullName
      .replace(/\s+/g, " ")
      .trim()
      .split(" ");

  // Skip "Md." or "Md"
  if (
    nameParts.length > 1 &&
    /^md\.?$/i.test(nameParts[0])
  ) {
    return nameParts[1];
  }

  return nameParts[0] || "Student";
}


function extractStudentNameFromExcel(
  workbook: XLSX.WorkBook
): string {
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

        // Example:
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
            return getFirstName(name);
          }

          if (values[i + 1]) {
            return getFirstName(
              values[i + 1]
            );
          }
        }

        // Example:
        // Name | Md. Mahir Hasan Shuvo
        if (
          /^name$/i.test(value) &&
          values[i + 1]
        ) {
          return getFirstName(
            values[i + 1]
          );
        }
      }
    }
  }

  return "Student";
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
  onStudentName?: (name: string) => void
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
const extractedName =
  extractStudentNameFromExcel(
    workbook
  );

onStudentName?.(extractedName);

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

  const [fileName, setFileName] =
    useState("");

  const [studentName, setStudentName] =
  useState("Student");
  

  const [classes, setClasses] =
    useState<ClassEntry[]>(
      []
    );

  const [loading, setLoading] =
    useState(false);

  const [error, setError] =
    useState("");

  const [editingIndex, setEditingIndex] =
    useState<number | null>(null);

  const [addingClass, setAddingClass] =
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
let pdfText = "";

if (
  extension === "pdf"
) {
  pdfText =
    await extractPdfText(
      file
    );

  setStudentName(
    extractStudentName(
      pdfText
    )
  );
}

let parsedClasses: ClassEntry[];

if (extension === "pdf") {
const pdfText =
  await extractPdfText(file);

console.log("===== PDF TEXT =====");
console.log(pdfText);
console.log("====================");

const extractedName =
  extractStudentName(pdfText);

console.log("===== EXTRACTED NAME =====");
console.log(extractedName);
console.log("==========================");

setStudentName(extractedName);

  parsedClasses =
    parsePdfText(pdfText);
} else {
parsedClasses =
  await parseFile(
    file,
    setStudentName
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

    <button
      className="add-class-btn"
      onClick={() => {
      setAddingClass(true);
     }}
      >
      + Add Class
      </button>

   </div>

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