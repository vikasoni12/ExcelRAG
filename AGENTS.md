Project: **ExcelRAG**
----------------------

**Excel-native RAG data preparation with Excel-DNA (net48)**

This project turns Excel into a **transparent, auditable RAG data prep pipeline** while respecting Excel’s recalculation engine, threading model, and UI constraints.

> ! This is **not** a normal UDF project.  
> Correct behavior depends on a small set of **non-obvious but proven patterns** documented below.

* * *

1. Core Design Philosophy (Read First)
--------------------------------------

### Excel is the orchestrator, not the runtime

Excel:

* Recalculates aggressively
* Replays formulas on open
* Has **no concept of state** between calls
* Runs UDFs on the **UI thread unless async/observable**

Therefore:

* **All persistent state lives outside Excel formulas** (registries keyed by caller + trigger)
* Excel formulas are _signals_, not commands
* Ingestion, chunking, and validation must be **idempotent**

* * *

2. Three Proven UDF Categories (Critical)
-----------------------------------------

### ① **Creation / Identity (Idempotent, Cached)**

Examples:

* `RAG.SOURCE_ID`
* `RAG.PIPELINE_ID`

**Rules**

* Must return the _same value_ when inputs are unchanged
* Must not recreate objects on every recalc
* Must be safe on workbook open

**Pattern**

* Caller-scoped cache (keyed by caller cell)
* Optional trigger to force regeneration
* Session-stable identity when needed

---

### ② **Triggered Work (Async, Fire-Once)**

Examples:

* `RAG.INGEST`
* `RAG.CHUNK`
* `RAG.EXPORT_JSONL`

**Rules**

* Must not run twice for the same trigger
* Must not block Excel UI
* Must tolerate recalculation storms

**Pattern**

* Trigger normalization (`TriggerKey`)
* LastTriggerKey guard
* `ExcelAsyncUtil.RunTask`
* Non-blocking `SemaphoreSlim`

---

### ③ **Observers (Push-based, No Recalc)**

Examples:

* `RAG.STATUS`
* `RAG.CHUNK_STATS`
* `RAG.VALIDATE_SUMMARY`

**Rules**

* Must never block
* Must not force workbook recalculation
* Must re-emit cached values when nothing changed
* Must tolerate spurious OnNext signals

**Pattern**

* `ExcelAsyncUtil.Observe`
* Central publish hub (e.g., `RagProgressHub`)
* Per-observable caching
* Push only on meaningful state transitions

* * *

3. Trigger Semantics (Exact)
----------------------------

### TriggerKey rules

* Scalars normalized
* Single-cell ranges coerced
* References dereferenced
* Stringified comparison

### Behavior

* **Same trigger → no-op**
* **Different trigger → new run**
* Returned value explains skip reason

* * *

4. Architecture Components (Target)
-----------------------------------

```
Excel
 ├─ Tables (Sources, Chunks, Metadata, QA)
 ├─ UDFs (RAG.INGEST, RAG.CHUNK, ...)
 │
Excel-DNA Add-in (.xll)
 ├─ Async-safe UDF execution
 ├─ Trigger-aware processing
 ├─ Registry + cache
 │
Managed .NET Layer
 ├─ Parsing, normalization
 ├─ Chunking + metadata
 ├─ Validation rules
 │
Export Layer
 ├─ JSONL, manifest
 └─ Provenance mapping
```

* * *

5. Code Generation Plan (Agent Guidance)
----------------------------------------

### Phase 1 — Foundation

* Create solution + add-in project targeting **net48**
* Add Excel-DNA references and build targets
* Implement **TriggerKey** utility
* Implement **caller-scoped cache** helper

### Phase 2 — Core UDFs

* `RAG.SOURCE_ID` (identity)
* `RAG.INGEST` (triggered work, async)
* `RAG.CHUNK` (triggered work, async)
* `RAG.METADATA` (deterministic transform)
* `RAG.STATUS` (observer)

### Phase 3 — Validation + Export

* `RAG.VALIDATE` + `RAG.VALIDATE_SUMMARY`
* `RAG.DEDUP`
* `RAG.EXPORT_JSONL` + `RAG.EXPORT_MANIFEST`

### Phase 4 — Ribbon

* Add a **RAG** ribbon tab
* Implement actions that **write tables** to the worksheet (no hidden state)
* Connect ribbon actions to existing UDFs with explicit trigger cells

* * *

6. Technical Deliverables (Required)
------------------------------------

* **Excel-DNA add-in** (.xll) with build + debug profile
* **UDF set** for ingest, chunk, metadata, validation, export
* **TriggerKey** and **registry/caches** (caller-scoped)
* **Observer hub** for progress + stats
* **Ribbon UI** with workflow actions
* **Example workbook** showing the end-to-end RAG prep flow
* **README** and **design notes** (this file)

* * *

7. Hard Rules (Do Not Deviate)
------------------------------

* **Never block Excel UI**
* **Never force workbook-wide recalculation**
* **All heavy work must be async**
* **Triggered work must be idempotent**
* **Observers are push-based, not recalc-driven**

* * *

8. Known Good / Known Bad
-------------------------

### ✔ Proven to work

* Trigger-guarded ingestion and chunking
* Push-based status updates
* Registry-backed identity functions
* Cache-first outputs with fallbacks

### ✖ Do NOT do

* Blocking `Wait()` on UI thread
* Recreating state per recalc
* Volatile identity for stable outputs
* Hidden background work without triggers

* * *

Final Note
----------

This architecture works **because it respects Excel’s rules instead of fighting them**.

If something feels “too careful,” it probably is — and it’s there because Excel crashes otherwise.
