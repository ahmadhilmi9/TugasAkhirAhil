# AGENTS.md — Tugas Akhir Scheduler

## Build & Run

- Requires **Java 23** and **Maven** (no wrapper).
- Build: `mvn clean package` — produces a fat JAR at `target/ta-1.0-SNAPSHOT.jar`.
- Run (dev): `mvn exec:java -Dexec.mainClass=org.example.Main`.
- Run (fat JAR): `java -jar target/ta-1.0-SNAPSHOT.jar`.

## Project Layout

Single-module Maven project.

| Path | Purpose |
|------|---------|
| `src/main/java/org/example/Main.java` | Entry point; reads `inputFilePath` from UI, processes Excel sheet index 3. |
| `src/main/java/org/example/SchedulerUI.java` | Swing-based UI; launched by `Main`. Exposes static file paths and scheduling parameters. |
| `output-jadwal.xlsx` | Default output file (generated at project root). |

## Dependencies

- **Apache POI 5.2.3** (`poi-ooxml`) — the only external dependency. Used for reading/writing `.xlsx` files.

## Testing

No test suite exists (`src/test/java` is empty, no JUnit dependency in `pom.xml`).

## Known Quirks

- `SchedulerUI.inputFilePath` must be set before calling `Main::runScheduler` — the UI file chooser populates it.
- Schedule parameters (number of classes, hours per day) are hardcoded as `static` fields in `SchedulerUI`; adjust them there if needed.
- The UI launches a Swing window; in headless environments the application will fail.
- Default output is written to `output-jadwal.xlsx` in the working directory.
