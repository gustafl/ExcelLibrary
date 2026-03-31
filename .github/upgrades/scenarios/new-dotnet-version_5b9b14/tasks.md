# ExcelLibrary .NET 10.0 Upgrade Tasks

## Overview

This document tracks the execution of the ExcelLibrary solution upgrade from .NET Framework 4.8 to .NET 10.0. Both the core library and test project will be upgraded simultaneously in a single atomic operation, followed by testing and validation.

**Progress**: 0/2 tasks complete (0%) ![0%](https://progress-bar.xyz/0)

---

## Tasks

### [▶] TASK-001: Atomic framework and dependency upgrade
**References**: Plan §Migration Strategy, Plan §Project-by-Project Plans, Plan §Package Update Reference, Plan §Breaking Changes Catalog

- [▶] (1) Convert ExcelLibrary project to SDK-style format per Plan §ExcelLibrary (use Microsoft.NET.Sdk, set TargetFramework to net10.0, disable ImplicitUsings and Nullable)
- [ ] (2) Convert ExcelLibrary.Tests project to SDK-style format per Plan §ExcelLibrary.Tests (use Microsoft.NET.Sdk, set TargetFramework to net10.0, enable IsTestProject, configure test asset copying for Input folder)
- [ ] (3) Add MSTest NuGet packages to ExcelLibrary.Tests per Plan §Package Update Reference (Microsoft.NET.Test.Sdk 17.*, MSTest.TestAdapter 3.*, MSTest.TestFramework 3.*)
- [ ] (4) Remove reference to Microsoft.VisualStudio.QualityTools.UnitTestFramework from ExcelLibrary.Tests
- [ ] (5) Restore all dependencies
- [ ] (6) All dependencies restored successfully (**Verify**)
- [ ] (7) Build solution and fix TimeSpan.FromSeconds breaking change in ExcelLibrary\Utilities.cs line 27 per Plan §Breaking Changes Catalog (validate the calculation produces reasonable TimeSpan values)
- [ ] (8) Solution builds with 0 errors (**Verify**)
- [ ] (9) Commit changes with message: "TASK-001: Complete atomic upgrade to .NET 10.0"

---

### [ ] TASK-002: Run full test suite and validate upgrade
**References**: Plan §Testing & Validation Strategy, Plan §Success Criteria

- [ ] (1) Run tests in ExcelLibrary.Tests project (verify all 4 test files are discovered: DefaultOptions.cs, LoadSheetsIsFalse.cs, IncludeHiddenIsTrue.cs, NumberFormats.cs)
- [ ] (2) Fix any test failures related to framework migration
- [ ] (3) Re-run tests after fixes
- [ ] (4) All tests pass with 0 failures (**Verify**)
- [ ] (5) Verify test input files (Input\test1.xlsx, Input\test2.xlsx) are copied to output directory (**Verify**)
- [ ] (6) Commit test fixes with message: "TASK-002: Complete testing and validation"

---
