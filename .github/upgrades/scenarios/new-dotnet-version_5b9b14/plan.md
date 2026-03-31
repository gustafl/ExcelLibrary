# .NET 10 Upgrade Plan

## Table of Contents

- [Executive Summary](#executive-summary)
- [Migration Strategy](#migration-strategy)
- [Detailed Dependency Analysis](#detailed-dependency-analysis)
- [Project-by-Project Plans](#project-by-project-plans)
  - [ExcelLibrary](#excellibrary)
  - [ExcelLibrary.Tests](#excellibrarytests)
- [Package Update Reference](#package-update-reference)
- [Breaking Changes Catalog](#breaking-changes-catalog)
- [Risk Management](#risk-management)
- [Testing & Validation Strategy](#testing--validation-strategy)
- [Complexity & Effort Assessment](#complexity--effort-assessment)
- [Source Control Strategy](#source-control-strategy)
- [Success Criteria](#success-criteria)

---

## Executive Summary

| Metric | Value |
|--------|-------|
| **Solution** | ExcelLibrary.sln |
| **Current Framework** | .NET Framework 4.8 |
| **Target Framework** | .NET 10.0 (LTS) |
| **Total Projects** | 2 |
| **Total Issues** | 5 (4 Mandatory, 1 Potential) |
| **Estimated Effort** | Low (~30 minutes) |
| **Risk Level** | Low |

### Scope Overview

This upgrade migrates the ExcelLibrary solution from .NET Framework 4.8 to .NET 10.0. The solution consists of:

1. **ExcelLibrary** - Core library for Excel file processing
2. **ExcelLibrary.Tests** - MSTest unit test project

### Key Changes Required

1. Convert both projects from legacy format to SDK-style project format
2. Update target framework from `net48` to `net10.0`
3. Migrate test framework from `Microsoft.VisualStudio.QualityTools.UnitTestFramework` to modern MSTest NuGet packages
4. Address `TimeSpan.FromSeconds` breaking change in .NET 10

---

## Migration Strategy

### Approach: Bottom-Up Sequential Migration

We will use a **bottom-up approach** based on project dependency order:

```
ExcelLibrary (core library - no dependencies)
    ↓
ExcelLibrary.Tests (depends on ExcelLibrary)
```

### Phases

| Phase | Description | Projects |
|-------|-------------|----------|
| **Phase 1** | SDK-style conversion & TFM update | ExcelLibrary |
| **Phase 2** | Fix breaking API changes | ExcelLibrary |
| **Phase 3** | SDK-style conversion & TFM update | ExcelLibrary.Tests |
| **Phase 4** | Test framework migration | ExcelLibrary.Tests |
| **Phase 5** | Validation | All |

---

## Detailed Dependency Analysis

### Project Dependency Graph

```
Level 0 (Foundation):
└── ExcelLibrary.csproj (no project dependencies)

Level 1 (Dependent):
└── ExcelLibrary.Tests.csproj
    └── References: ExcelLibrary
```

### ExcelLibrary Dependencies

| Dependency | Type | .NET 10 Status |
|------------|------|----------------|
| System | Framework | ✅ Built-in |
| System.Core | Framework | ✅ Built-in |
| System.IO.Compression | Framework | ✅ Built-in |
| System.IO.Compression.FileSystem | Framework | ⚠️ Removed - functionality in System.IO.Compression |
| System.Xml.Linq | Framework | ✅ Built-in |
| System.Data.DataSetExtensions | Framework | ✅ Built-in |
| Microsoft.CSharp | Framework | ✅ Built-in |
| System.Data | Framework | ✅ Built-in |
| System.Xml | Framework | ✅ Built-in |

### ExcelLibrary.Tests Dependencies

| Dependency | Type | .NET 10 Status |
|------------|------|----------------|
| System | Framework | ✅ Built-in |
| Microsoft.VisualStudio.QualityTools.UnitTestFramework | GAC Assembly | ❌ Replace with MSTest NuGet |
| ExcelLibrary | Project Reference | ✅ Will be upgraded |

---

## Project-by-Project Plans

### ExcelLibrary

**Path:** `ExcelLibrary\ExcelLibrary.csproj`  
**Type:** Class Library  
**Current TFM:** net48  
**Target TFM:** net10.0

#### Actions Required

1. **Convert to SDK-style project format**
   - Replace entire project file content with SDK-style format
   - Use `Microsoft.NET.Sdk` SDK
   - Remove explicit file includes (SDK-style uses automatic inclusion)
   - Remove legacy property groups and imports

2. **Update target framework**
   - Change `<TargetFramework>` to `net10.0`

3. **Remove obsolete framework references**
   - SDK-style projects automatically reference required BCL assemblies
   - `System.IO.Compression.FileSystem` functionality is included in `System.IO.Compression`

4. **Fix breaking API change** (see Breaking Changes Catalog)
   - Update `TimeSpan.FromSeconds` call in `Utilities.cs`

#### Target Project File

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <RootNamespace>ExcelLibrary</RootNamespace>
    <AssemblyName>ExcelLibrary</AssemblyName>
    <ImplicitUsings>disable</ImplicitUsings>
    <Nullable>disable</Nullable>
  </PropertyGroup>

</Project>
```

### ExcelLibrary.Tests

**Path:** `ExcelLibrary.Tests\ExcelLibrary.Tests.csproj`  
**Type:** MSTest Unit Test Project  
**Current TFM:** net48  
**Target TFM:** net10.0

#### Actions Required

1. **Convert to SDK-style project format**
   - Replace entire project file content with SDK-style format
   - Use `Microsoft.NET.Sdk` SDK
   - Enable `IsTestProject` property

2. **Update target framework**
   - Change `<TargetFramework>` to `net10.0`

3. **Migrate test framework**
   - Remove reference to `Microsoft.VisualStudio.QualityTools.UnitTestFramework`
   - Add NuGet packages:
     - `Microsoft.NET.Test.Sdk`
     - `MSTest.TestAdapter`
     - `MSTest.TestFramework`

4. **Configure test assets**
   - Ensure test input files (`Input\test1.xlsx`, `Input\test2.xlsx`) are copied to output

#### Target Project File

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <RootNamespace>ExcelLibrary.Tests</RootNamespace>
    <AssemblyName>ExcelLibrary.Tests</AssemblyName>
    <IsPackable>false</IsPackable>
    <IsTestProject>true</IsTestProject>
    <ImplicitUsings>disable</ImplicitUsings>
    <Nullable>disable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.*" />
    <PackageReference Include="MSTest.TestAdapter" Version="3.*" />
    <PackageReference Include="MSTest.TestFramework" Version="3.*" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\ExcelLibrary\ExcelLibrary.csproj" />
  </ItemGroup>

  <ItemGroup>
    <None Include="Input\**\*" CopyToOutputDirectory="PreserveNewest" />
  </ItemGroup>

</Project>
```

---

## Package Update Reference

### Packages to Add

| Project | Package | Version | Purpose |
|---------|---------|---------|---------|
| ExcelLibrary.Tests | Microsoft.NET.Test.Sdk | 17.* | Test host and execution |
| ExcelLibrary.Tests | MSTest.TestAdapter | 3.* | Test discovery and execution |
| ExcelLibrary.Tests | MSTest.TestFramework | 3.* | Test attributes and assertions |

### Packages/References to Remove

| Project | Reference | Reason |
|---------|-----------|--------|
| ExcelLibrary.Tests | Microsoft.VisualStudio.QualityTools.UnitTestFramework | Not available in .NET 10; replaced by MSTest NuGet packages |

---

## Breaking Changes Catalog

### Api.0002: TimeSpan.FromSeconds Breaking Change

**Severity:** Potential  
**Location:** `ExcelLibrary\Utilities.cs`, Line 27  
**Affected Code:**
```csharp
TimeSpan span = TimeSpan.FromSeconds(seconds);
```

**Description:**  
In .NET 10, `TimeSpan.FromSeconds(double)` has improved precision and stricter validation. Large values that previously worked may now throw `OverflowException` if they exceed `TimeSpan.MaxValue`.

**Resolution:**  
The current code calculates Excel time values, which should produce reasonable TimeSpan values (within 24 hours). However, to ensure robustness:

1. **Option A (Recommended):** Keep the code as-is if the input is always valid Excel time fractions (0.0 to 1.0 representing 24-hour time)
2. **Option B:** Add validation to handle edge cases:
   ```csharp
   public static string ConvertTime(string excelTime)
   {
       double time = double.Parse(excelTime, CultureInfo.GetCultureInfo("en-us"));
       double second = 1 / 86400d;
       double seconds = time / second;

       // Clamp to valid TimeSpan range
       if (seconds > TimeSpan.MaxValue.TotalSeconds)
           seconds = TimeSpan.MaxValue.TotalSeconds;

       TimeSpan span = TimeSpan.FromSeconds(seconds);
       return span.ToString();
   }
   ```

**Recommendation:** Validate at runtime. The current usage with Excel time fractions should work correctly.

---

## Risk Management

### Risk Assessment

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| TimeSpan overflow for edge cases | Low | Low | Add validation if tests fail |
| Test discovery issues | Low | Medium | Verify test packages are correctly configured |
| Build errors from implicit usings | Low | Low | Disabled implicit usings to maintain compatibility |

### Rollback Strategy

1. All changes are on branch `upgrade-to-NET10`
2. Original code remains intact on `master` branch
3. If issues arise, switch back to `master` branch

---

## Testing & Validation Strategy

### Validation Checkpoints

| Checkpoint | Criteria | Method |
|------------|----------|--------|
| **After Phase 1** | ExcelLibrary builds successfully | `dotnet build` |
| **After Phase 2** | No compilation errors from API changes | `dotnet build` |
| **After Phase 3** | ExcelLibrary.Tests builds successfully | `dotnet build` |
| **After Phase 4** | All tests pass | `dotnet test` |
| **Final** | Solution builds and all tests pass | Full build + test |

### Test Execution Plan

1. Run existing unit tests after migration
2. Verify test discovery in Visual Studio Test Explorer
3. Confirm all 4 test files are discovered:
   - `DefaultOptions.cs`
   - `LoadSheetsIsFalse.cs`
   - `IncludeHiddenIsTrue.cs`
   - `NumberFormats.cs`

---

## Complexity & Effort Assessment

### Overall Assessment: **Low Complexity**

| Factor | Assessment | Notes |
|--------|------------|-------|
| Project Count | Simple (2 projects) | Minimal coordination needed |
| Dependencies | Simple | All framework references, no complex NuGet packages |
| Breaking Changes | Minimal (1 issue) | TimeSpan change is low risk |
| Test Framework | Moderate | Requires package migration |
| Code Changes | Minimal | Likely no code changes needed |

### Estimated Timeline

| Task | Estimated Time |
|------|----------------|
| Convert ExcelLibrary to SDK-style | 5 minutes |
| Convert ExcelLibrary.Tests to SDK-style | 10 minutes |
| Validate builds | 5 minutes |
| Run and verify tests | 5 minutes |
| **Total** | **~25-30 minutes** |

---

## Source Control Strategy

### Branch Strategy

| Branch | Purpose |
|--------|---------|
| `master` | Original .NET Framework 4.8 code (unchanged) |
| `upgrade-to-NET10` | All upgrade changes |

### Commit Strategy

| Commit | Content |
|--------|---------|
| 1 | Convert ExcelLibrary to SDK-style, target net10.0 |
| 2 | Convert ExcelLibrary.Tests to SDK-style, target net10.0, migrate test framework |
| 3 | Fix any breaking changes (if needed) |
| 4 | Final validation and cleanup |

---

## Success Criteria

### Definition of Done

- [ ] Both projects successfully converted to SDK-style format
- [ ] Both projects target `net10.0`
- [ ] Solution builds without errors
- [ ] All unit tests pass
- [ ] Test input files are correctly copied to output directory
- [ ] No runtime errors in test execution

### Verification Commands

```bash
# Build solution
dotnet build ExcelLibrary.sln

# Run tests
dotnet test ExcelLibrary.sln

# Verify target framework
dotnet --list-sdks
```
