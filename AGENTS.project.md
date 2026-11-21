## Project specifics
- Access add-in "DBMS Connection Wizard": configures a DBMS connection and imports the
  ACLib DbConnection framework into the target database.
- `source\` (msaccess-vcs-addin export) is the single source of truth; binary files
  (`.accdb`, `.accda`) are not edited as source and are not committed.
- The add-in file is built from `source\` with msaccess-vcs-build
  (GitHub action or local `Build.ps1`); do not run a build without approval.
- `doc\` holds the interface documentation generated with Natural Docs (`doc\GenerateDoc.cmd`);
  do not edit generated output by hand.
- `.accdb`/`.accda` files in the root (e.g. `Test.accdb`, `TestAddIn.accda`) are local test or build files.
