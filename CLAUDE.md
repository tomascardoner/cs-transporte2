# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

CS-Transporte is a Windows desktop line-of-business application (VB6) for managing a minibus/passenger-transport
company: routes and schedules, trips (Viaje) and their passenger/commission detail (ViajeDetalle), fleet and
maintenance, driver assignments, price lists, client accounts (Cuenta Corriente), web reservations, and Crystal
Reports reporting. All UI text, business terms, and most identifiers are in Spanish; the business domain is
Argentina (AFIP-related helpers exist in the shared framework).

The same codebase/schema is deployed to several separate SQL Server databases, one per client company (e.g.
`CSTransporte_LobosBus`, `CSTransporte_DelSurBus`, `CSTransporte_TuBus`) — see `Stored procedures/Generate.bat` and
the `[Database]` section of `CSTransporte.ini`. Some report/stored-procedure files are company-specific (files
named `... - LobosBus.sql`, `... - DelSurBus.sql`).

## Repository layout

- `VB Project/` — the VB6 application: `CSTransporte.vbp` project file, `.frm`/`.frx` forms, `.cls` classes,
  `.bas` modules. This is the only buildable code in the repo.
- `Stored procedures/` — one `.sql` file per entity (e.g. `Viaje.sql`, `Persona.sql`, `Ruta.sql`) containing the SQL
  Server stored procedures that back that entity's class. `Generate.bat` runs them all against a target database
  via `sqlcmd`; company-specific one-off generate scripts also exist for individual environments.
- `Reports/` — Crystal Reports `.rpt` (and a couple `.xls`/`.ttx`) report definitions, loaded at runtime by
  `Reporte.cls` and shown via `ReportViewer.frm`.
- `Version Launcher/` — a small standalone C# WinForms app (`CSLauncher`, .NET, `CSLauncher.sln`) that launches
  `CSTransporte.exe` with `CONFIG=<name>`, where `<name>` is the launcher's own command-line argument (an INI file
  base name). Used so a desktop shortcut can start the app against a specific company's config without hardcoding
  paths in the shortcut itself.

## Critical external dependency: `../framework-vb6`

`CSTransporte.vbp` references a large number of modules/classes by relative path into a **sibling repository**,
`../framework-vb6` (e.g. `Module=CSM_Session; ..\..\framework-vb6\CSM_Session.bas`,
`Class=CSC_Database_ADO_SQL; ..\..\framework-vb6\CSC_Database_ADO_SQL.cls`). This shared framework provides generic
infrastructure — ADO database access, DES encryption, INI/registry helpers, application logging, form
position/sizing, common dialogs, AFIP web services, telephony (TAPI) — that is reused across multiple other
Cardoner Sistemas products in sibling directories (`cs-gestion`, `cs-cereal`, `cs-agroindustria`, etc.). This repo
must be checked out as a sibling of `framework-vb6` for the project to open/compile, and **changes to files under
`framework-vb6` affect every product that references it**, not just this one — treat edits there as higher-risk
than edits inside `VB Project/`.

## Building / running

There is no CLI build. This is a VB6 (pre-.NET) project:

- Open `VB Project/CSTransporte.vbp` in the VB6 IDE (32-bit, requires VB6 SP6) to edit, run, or compile.
- The IDE reference list in the `.vbp` requires these to be installed/registered on the machine: Crystal Reports
  11.5 runtime (`craxdrt.dll`, `crdb_ado.dll`), ComponentOne `xadb8.ocx`, `todg8.ocx`, MS ADO 6.1, `mscomctl.OCX`,
  `COMCT332.OCX`, `MSCOMCT2.OCX`, `MSCOMM32.OCX` (TAPI), `MSDATLST.OCX`.
- `Startup="Sub Main"` (in `Startup.bas`) is the actual entry point, not any form's `Load` event.
- Runtime configuration comes from an INI file (default `CSTransporte.ini`, see `[Database]`, `[Logs]`, `[Reports]`,
  `[Interface]`, `[System]`). The active config can be switched via the `CONFIG=<name>` command-line argument
  (loads `<name>.ini` instead), which is exactly what `Version Launcher` uses to run multiple company configs from
  one install.
- There is no automated test suite; verification is manual, through the compiled app against a target database.

## Database / stored procedures

- SQL Server backend, accessed only through stored procedures (never ad-hoc SQL from the app, aside from a few
  read-only SELECTs built dynamically for lists/permission filters).
- Naming convention: `sp_<Entity>_<Data|Insert|Update|Delete>` (e.g. `sp_Viaje_Data`, `sp_Viaje_Insert`). Each
  script in `Stored procedures/` is idempotent — it drops-if-exists then recreates every procedure — so re-running
  `Generate.bat` (or a single entity's `.sql` file via `sqlcmd`) against a database is the normal way to push schema
  logic changes.
- `Stored procedures/Report.sql` / `Report - ViajePlanilla.sql` back the Crystal Reports parameter/data procedures;
  report-specific per-company scripts under `Stored procedures/Reporte - *.sql` are for the web-reservation reports
  in `Reports/`.

## Application architecture (VB Project)

**Entity classes (`.cls`).** Each business entity (`Viaje`, `Persona`, `Ruta`, `Vehiculo`, `CuentaCorriente`, ...)
is a class module following the same shape:
- `Private m<Field>` backing fields plus `Property Get`/`Property Let` pairs; every `Let` sets `mIsDirty = True`
  when the value actually changes.
- State flags: `mIsNew`, `mIsCopy`, `mIsDirty`, `mNoMatch` (and `mNoMatchRaiseError` to control whether a missing
  record is an error or an expected "not found").
- `OpenRecordset()` → opens the entity's `sp_<Entity>_Data` recordset; `Load()` calls it and populates fields (also
  snapshotting `...Original` copies of key fields used later to detect what changed); `Update()` inserts or updates
  via the `_Insert`/`_Update` procedures; `Delete()`; some entities add a `Copy()`.
  Errors follow a standard pattern: `If pTrapErrors Then On Error GoTo ErrorHandler`, then on the label call
  `ShowErrorMessage "Classes.<Entity>.<Method>", "<Spanish message>"` (from the shared framework's `CSM_Error.bas`).
- Larger entities (`Viaje`, `ViajeDetalle`) add domain logic beyond CRUD in the same class — seat assignment
  (`Asiento_Asignar*`), account-ledger posting (`CuentaCorriente_Asignar*`), fixed-reservation generation
  (`GenerarReservasFijas`), overdue-status checks, etc.

**Forms (`.frm`).** Each entity typically has a list/search form (`Viaje.frm`) and a separate detail/edit form
named `<Entity>Propiedad.frm` (e.g. `ViajePropiedad.frm`, `PersonaRutaPropiedad.frm`) — the ABM (Alta/Baja/
Modificación = Create/Delete/Update) convention used throughout. Group/multi-select variants exist too
(`ViajePropiedadMultiple.frm`, `ComisionAsistenciaMultiple.frm`).

**Permissions.** `Permiso_Module.bas` defines every permission as a `PERMISO_*` string constant (e.g.
`PERMISO_VIAJE_MODIFY`, `PERMISO_VIAJE_DETALLE_TRANSFER_ALLOWOUTOFRANGE`). At login, `CPermiso.cls` loads the
current user-group's granted `Permiso` objects into a collection and exposes `GotPermission(...)` checks used to
show/hide UI and to build row-level-security SQL fragments dynamically (e.g. `RutaWhere`, `CuentaCorrienteCajaWhere`,
`CuentaCorrienteGrupoWhere`) that get appended to list queries so a user only sees rows they're permitted to see.
When adding a new permission, add the constant in `Permiso_Module.bas`, wire the check into the relevant form(s),
and add the row in `Stored procedures/Permissions.sql`.

**Global/ambient state.** `Startup.bas` module-level `Public` variables (`pParametro`, `pDatabase`, `pUsuario`,
`pCPermiso`, `pSucursal`, `pTelephony`, `pRegionalSettings`, ...) are set once in `Sub Main` and then read directly
from anywhere in the app — there is no dependency injection; forms and classes reach into these globals rather
than receiving them as parameters.

**Multi-user refresh ("Semáforo").** Because multiple workstations share one database with no push notifications,
the app polls: `RefreshList_Module.bas` defines a `MODULE_*` code per entity, and a timer
(`frmMDI.tmrMulti`) periodically checks the `Semaforo`/`SemaforoGeneral` table (via `precSemaforoGeneral`,
`pCSemaforoGeneral`) for a changed counter per module and triggers that module's list to refresh in the open forms.
When an entity's data changes, its counter must be bumped so other open sessions notice.

**Cross-cutting helpers.** `CSTransporte_SDK.bas` holds shared, form-aware helpers used from many places (open a
trip + its detail together, resolve "the current trip" for a vehicle within a configurable time window, set the
caller-ID/last-person combo on the MDI, write to `LogAccion`, substitute `|@Token|`-style placeholders in outgoing
message text). `Constants.bas` holds shared domain string/code constants (entity type codes, `Viaje` estado codes,
reservation type codes, etc.) — prefer reusing these over hardcoding the underlying string literals.

**Reports.** `Reporte.cls`/`ReporteParametro.cls` load a `.rpt` from `Reports/`, bind parameters (typed via
`REPORTE_PARAMETRO_TIPO_*` constants in `Constants.bas`), and display it through `ReportViewer.frm` using the
Crystal Reports ActiveX runtime (CRAXDRT) — `pCRAXDRTApplication` is initialized once in `Startup.bas` and logged
off in `TerminateApplication`.

## Conventions to follow when editing

- Match the existing entity-class shape (backing field + Get/Let + dirty tracking) rather than introducing a new
  pattern for a single entity.
- Keep stored-procedure scripts idempotent (`IF EXISTS ... DROP PROCEDURE` before `CREATE PROCEDURE`) and named
  `sp_<Entity>_<Data|Insert|Update|Delete>` to match the existing set.
- New permissions go through `Permiso_Module.bas` + `Permissions.sql`, not ad-hoc boolean flags.
- Commit messages in this repo are short, plain Spanish, present/past-tense descriptions of the change (e.g. "Se
  agregó un nuevo permiso para viajes finalizados", "Bug fix: No asignaba IDViaje a las reservas fijas..."); follow
  that style rather than Conventional Commits.
