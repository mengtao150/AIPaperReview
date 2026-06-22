# Narrative Review MVP Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a runnable first version that imports journal checklists, accepts DOCX/PDF manuscripts, executes traceable Narrative Review checks, and lets a reviewer amend and export the report.

**Architecture:** Use a Python FastAPI modular monolith with SQLite for persistent metadata and local filesystem storage for uploaded files. Convert every manuscript into a `DocumentGraph` with stable anchors, run deterministic rules locally, and isolate semantic review behind an evaluator interface with a safe heuristic fallback and an optional LLM adapter. A small server-rendered HTML workspace uses the same JSON APIs, so the MVP is usable without a separate frontend build.

**Tech Stack:** Python 3.12, FastAPI, SQLAlchemy, Pydantic, Jinja2, SQLite, pytest, python-docx, pypdf, openpyxl, Uvicorn.

---

## File structure

```text
AIPaperReview/
  app/
    main.py                    # FastAPI app, routes, templates/static mounting
    config.py                  # settings and upload paths
    db.py                      # SQLAlchemy engine/session setup
    models.py                  # persistent entities
    schemas.py                 # API request/response models
    services/
      rule_import.py           # DOCX/PDF/XLSX checklist extraction
      document_parse.py        # DOCX/PDF -> DocumentGraph
      rules.py                 # published Narrative Review baseline rules
      deterministic_review.py  # count/structure/cross-reference checks
      semantic_review.py       # evaluator protocol and heuristic evaluator
      review_service.py        # orchestration, evidence validation, persistence
      report_service.py        # bilingual Markdown report generation
    templates/
      index.html               # upload and review workbench
    static/
      app.js                   # API client and result rendering
      app.css                  # accessible MVP styling
  tests/
    conftest.py
    fixtures/
      narrative_complete.docx
      narrative_incomplete.docx
      checklist.xlsx
    test_rule_import.py
    test_document_parse.py
    test_deterministic_review.py
    test_review_api.py
  requirements.txt
  README.md
```

### Task 1: Create the runnable application foundation

**Files:**
- Create: `requirements.txt`
- Create: `app/__init__.py`
- Create: `app/config.py`
- Create: `app/db.py`
- Create: `app/main.py`
- Create: `tests/conftest.py`
- Create: `tests/test_health.py`

- [ ] **Step 1: Write the failing health endpoint test**

```python
from fastapi.testclient import TestClient
from app.main import create_app

def test_health_reports_ready(tmp_path, monkeypatch):
    monkeypatch.setenv("APP_DATA_DIR", str(tmp_path))
    client = TestClient(create_app())
    assert client.get("/api/health").json() == {"status": "ok"}
```

- [ ] **Step 2: Run the test to confirm it fails**

Run: `pytest tests/test_health.py -q`

Expected: FAIL because `app.main` does not exist.

- [ ] **Step 3: Add minimal dependency and app bootstrap code**

`requirements.txt` must pin FastAPI, Uvicorn, SQLAlchemy, Pydantic, Jinja2, python-docx, pypdf, openpyxl, pytest and httpx. `create_app()` must create upload/report directories and register `GET /api/health`, returning `{"status": "ok"}`.

```python
def create_app() -> FastAPI:
    app = FastAPI(title="AI Paper Review")
    @app.get("/api/health")
    def health() -> dict[str, str]:
        return {"status": "ok"}
    return app
```

- [ ] **Step 4: Run the health test**

Run: `pytest tests/test_health.py -q`

Expected: PASS, `1 passed`.

- [ ] **Step 5: Commit the foundation**

```bash
git add requirements.txt app tests/test_health.py tests/conftest.py
git commit -m "feat: scaffold review service"
```

### Task 2: Persist templates, manuscripts, review runs and findings

**Files:**
- Create: `app/models.py`
- Create: `app/schemas.py`
- Modify: `app/db.py`
- Create: `tests/test_models.py`

- [ ] **Step 1: Write the failing persistence test**

```python
def test_review_run_keeps_template_version_and_findings(session):
    template = Template(name="AME Narrative", article_type="narrative_review", version="1.0.0")
    manuscript = Manuscript(filename="paper.docx", sha256="a" * 64)
    run = ReviewRun(template=template, manuscript=manuscript, report_language="zh")
    run.findings.append(Finding(rule_id="methods.search_date", status="missing", confidence=0.9))
    session.add(run)
    session.commit()
    assert session.get(ReviewRun, run.id).findings[0].status == "missing"
```

- [ ] **Step 2: Run the model test to confirm it fails**

Run: `pytest tests/test_models.py -q`

Expected: FAIL because the entities are undefined.

- [ ] **Step 3: Implement focused SQLAlchemy entities**

Create `Template`, `Rule`, `Manuscript`, `ReviewRun`, `Finding`, and `FindingDecision` entities. Include immutable `template_version`, JSON fields for rule configuration/evidence, status enums as strings, `created_at` timestamps, and foreign keys. Add `create_all()` initialization in `db.py`; `tests/conftest.py` supplies an isolated SQLite session.

- [ ] **Step 4: Run the model test**

Run: `pytest tests/test_models.py -q`

Expected: PASS, `1 passed`.

- [ ] **Step 5: Commit persistence**

```bash
git add app/db.py app/models.py app/schemas.py tests/conftest.py tests/test_models.py
git commit -m "feat: persist review domain records"
```

### Task 3: Import rule checklists and publish a versioned template

**Files:**
- Create: `app/services/rule_import.py`
- Create: `app/services/rules.py`
- Modify: `app/main.py`
- Create: `tests/test_rule_import.py`

- [ ] **Step 1: Write failing rule-import tests**

```python
def test_xlsx_checklist_becomes_editable_candidate_rules(checklist_file):
    rules = import_checklist(checklist_file)
    assert rules[0].source["sheet"] == "Narrative"
    assert any("search" in rule.title_en.lower() for rule in rules)

def test_baseline_has_precise_search_date_rule():
    rules = narrative_review_baseline()
    rule = next(rule for rule in rules if rule.id == "methods.search_date")
    assert rule.rule_type == "semantic_presence"
    assert rule.required is True
```

- [ ] **Step 2: Run the import tests to confirm they fail**

Run: `pytest tests/test_rule_import.py -q`

Expected: FAIL because the services are missing.

- [ ] **Step 3: Implement rule extraction and baseline rules**

Implement file-type dispatch: DOCX paragraphs/tables, PDF page text, and Excel rows. Emit candidates with source locations, never published rules. Implement `narrative_review_baseline()` with at least: title focus/type, abstract components, keyword count, introduction background/rationale/objective, methods search date/range/database/strategy/inclusion criteria/screening, body limitations, conclusion significance, word limit, reference and figure checks. Add `POST /api/templates/import` and `POST /api/templates/baseline` endpoints, with an explicit `publish` request that increments template version.

- [ ] **Step 4: Run the import tests**

Run: `pytest tests/test_rule_import.py -q`

Expected: PASS.

- [ ] **Step 5: Commit template functionality**

```bash
git add app/services/rule_import.py app/services/rules.py app/main.py tests/test_rule_import.py tests/fixtures/checklist.xlsx
git commit -m "feat: import and version review templates"
```

### Task 4: Parse DOCX and PDF files into evidence-addressable document graphs

**Files:**
- Create: `app/services/document_parse.py`
- Create: `tests/test_document_parse.py`
- Create: `tests/fixtures/narrative_complete.docx`
- Create: `tests/fixtures/narrative_incomplete.docx`

- [ ] **Step 1: Write failing parser tests**

```python
def test_docx_parser_assigns_sections_and_paragraph_anchors(complete_docx):
    graph = parse_document(complete_docx)
    assert graph.section("Methods") is not None
    assert graph.section("Methods").blocks[0].anchor.startswith("docx:")

def test_pdf_parser_preserves_page_anchors(pdf_from_complete_docx):
    graph = parse_document(pdf_from_complete_docx)
    assert all(block.anchor.startswith("pdf:p") for block in graph.blocks)
```

- [ ] **Step 2: Run the parser tests to confirm they fail**

Run: `pytest tests/test_document_parse.py -q`

Expected: FAIL because `parse_document` is undefined.

- [ ] **Step 3: Implement DocumentGraph and parsers**

Define `DocumentGraph`, `Section`, `Block`, `TableBlock`, and `Asset` dataclasses. DOCX parsing must use heading styles and normalized heading names; PDF parsing must use `pypdf.PdfReader`, produce one text block per page with `pdf:p{page}:b{index}` anchors, and refuse scanned PDFs that have no extractable text with an actionable OCR error. Extract keyword candidates, word count, heading tree, figures/tables and reference-section blocks.

- [ ] **Step 4: Run the parser tests**

Run: `pytest tests/test_document_parse.py -q`

Expected: PASS.

- [ ] **Step 5: Commit document parsing**

```bash
git add app/services/document_parse.py tests/test_document_parse.py tests/fixtures
git commit -m "feat: parse manuscripts with traceable anchors"
```

### Task 5: Implement deterministic Narrative Review checks

**Files:**
- Create: `app/services/deterministic_review.py`
- Create: `tests/test_deterministic_review.py`

- [ ] **Step 1: Write failing deterministic-review tests**

```python
def test_incomplete_manuscript_reports_missing_keyword_count_and_methods(incomplete_graph):
    findings = run_deterministic_rules(incomplete_graph, narrative_review_baseline())
    by_id = {finding.rule_id: finding for finding in findings}
    assert by_id["keywords.count"].status == "missing"
    assert by_id["methods.section"].status == "missing"

def test_complete_manuscript_passes_word_limit_and_structure(complete_graph):
    findings = run_deterministic_rules(complete_graph, narrative_review_baseline())
    by_id = {finding.rule_id: finding for finding in findings}
    assert by_id["word_count.limit"].status == "present"
    assert by_id["conclusion.section"].status == "present"
```

- [ ] **Step 2: Run the deterministic tests to confirm they fail**

Run: `pytest tests/test_deterministic_review.py -q`

Expected: FAIL because the rule engine is missing.

- [ ] **Step 3: Implement only evidence-backed deterministic checks**

For `count`, `structure`, and `cross_reference` rules, return a `Finding` with exactly one of `present`, `partial`, `missing`, `uncertain`, or `not_applicable`, a source anchor where present, and a concrete suggestion where missing. Include keyword count, word limit, required sections, “Main body” heading prohibition, exact date pattern presence, table/figure textual references, and a reference-section presence check. Never report copyright or visual quality as passing; return `uncertain` and require human review.

- [ ] **Step 4: Run the deterministic tests**

Run: `pytest tests/test_deterministic_review.py -q`

Expected: PASS.

- [ ] **Step 5: Commit deterministic checks**

```bash
git add app/services/deterministic_review.py tests/test_deterministic_review.py
git commit -m "feat: add evidence-backed structural review"
```

### Task 6: Add semantic evaluation with a safe fallback

**Files:**
- Create: `app/services/semantic_review.py`
- Create: `tests/test_semantic_review.py`
- Modify: `app/config.py`

- [ ] **Step 1: Write failing semantic-evaluator tests**

```python
def test_heuristic_evaluator_marks_missing_rationale_without_evidence(incomplete_graph, rationale_rule):
    finding = HeuristicEvaluator().evaluate(rationale_rule, incomplete_graph)
    assert finding.status == "missing"
    assert finding.confidence < 0.8

def test_semantic_finding_never_contains_unknown_anchor(complete_graph, objective_rule):
    finding = HeuristicEvaluator().evaluate(objective_rule, complete_graph)
    assert all(anchor in complete_graph.anchor_ids for anchor in finding.evidence_anchors)
```

- [ ] **Step 2: Run semantic tests to confirm they fail**

Run: `pytest tests/test_semantic_review.py -q`

Expected: FAIL because the evaluator is missing.

- [ ] **Step 3: Implement evaluator protocol and heuristic MVP**

Define `SemanticEvaluator.evaluate(rule, graph) -> Finding`. The default evaluator searches normalized rule-specific cue phrases in the configured target sections and returns `present`, `partial`, `missing`, or `uncertain` with graph anchors. Add an `LLM_EVALUATOR_ENABLED` setting but do not make the MVP depend on credentials; the optional adapter is only called after the exact rule, evidence chunks, JSON schema and allowed anchors are passed to it. Validate any adapter result server-side and downgrade invalid anchors or unsupported assertions to `uncertain`.

- [ ] **Step 4: Run semantic tests**

Run: `pytest tests/test_semantic_review.py -q`

Expected: PASS.

- [ ] **Step 5: Commit semantic evaluation**

```bash
git add app/config.py app/services/semantic_review.py tests/test_semantic_review.py
git commit -m "feat: add constrained semantic review"
```

### Task 7: Orchestrate review runs and expose a reviewer API

**Files:**
- Create: `app/services/review_service.py`
- Modify: `app/main.py`
- Create: `tests/test_review_api.py`

- [ ] **Step 1: Write failing end-to-end API tests**

```python
def test_reviewer_can_create_review_and_override_a_finding(client, complete_docx):
    template = client.post("/api/templates/baseline").json()
    upload = client.post("/api/manuscripts", files={"file": complete_docx})
    review = client.post("/api/reviews", json={
        "manuscript_id": upload.json()["id"],
        "template_id": template["id"],
        "language": "en",
    }).json()
    finding = client.get(f"/api/reviews/{review['id']}").json()["findings"][0]
    updated = client.patch(f"/api/findings/{finding['id']}", json={
        "decision": "edited_accept", "status": "partial", "comment": "Reviewer clarification"
    })
    assert updated.json()["final_status"] == "partial"
```

- [ ] **Step 2: Run API test to confirm it fails**

Run: `pytest tests/test_review_api.py -q`

Expected: FAIL because manuscript and review endpoints are absent.

- [ ] **Step 3: Implement review orchestration and endpoints**

Implement `POST /api/manuscripts`, `POST /api/reviews`, `GET /api/reviews/{id}`, `GET /api/reviews/{id}/findings`, and `PATCH /api/findings/{id}`. A run must bind a published template version, parse the manuscript, execute deterministic and semantic rules, validate evidence anchors, persist initial AI findings, and prevent a rerun from overwriting reviewer decisions. Return normalized error bodies for invalid formats, unavailable templates, parsing errors, and invalid state transitions.

- [ ] **Step 4: Run API tests**

Run: `pytest tests/test_review_api.py -q`

Expected: PASS.

- [ ] **Step 5: Commit review API**

```bash
git add app/main.py app/services/review_service.py tests/test_review_api.py
git commit -m "feat: run and review manuscript checks"
```

### Task 8: Provide bilingual reports and a usable browser workbench

**Files:**
- Create: `app/services/report_service.py`
- Create: `app/templates/index.html`
- Create: `app/static/app.js`
- Create: `app/static/app.css`
- Modify: `app/main.py`
- Create: `tests/test_report_service.py`

- [ ] **Step 1: Write failing report test**

```python
def test_chinese_report_uses_reviewer_final_status(review_with_override):
    report = render_markdown_report(review_with_override, language="zh")
    assert "结构审稿" in report
    assert "部分到位" in report
    assert "AI 初审经人工复核" in report
```

- [ ] **Step 2: Run report test to confirm it fails**

Run: `pytest tests/test_report_service.py -q`

Expected: FAIL because the report renderer is missing.

- [ ] **Step 3: Implement reports and the minimal workbench**

Generate Markdown reports containing template version, overview counts, structural findings, content findings, uncertain/manual-review findings, evidence anchors, reviewer comments and the human-review marker. Map statuses and headings through a single Chinese/English dictionary. Add `GET /api/reviews/{id}/report?language=zh|en` and `GET /` UI: file upload, template creation, start review, status filters, side-by-side finding/evidence view, and accept/edit/reject controls. Escape all manuscript-derived text before inserting it into HTML.

- [ ] **Step 4: Run report and full test suites**

Run: `pytest tests/test_report_service.py -q && pytest -q`

Expected: PASS with no failed tests.

- [ ] **Step 5: Commit the workbench**

```bash
git add app/services/report_service.py app/templates app/static app/main.py tests/test_report_service.py
git commit -m "feat: add bilingual review reports and workbench"
```

### Task 9: Document local operation and validate the production path

**Files:**
- Create: `README.md`
- Modify: `app/main.py`
- Create: `tests/test_security.py`

- [ ] **Step 1: Write failing upload-hardening test**

```python
def test_upload_rejects_non_manuscript_extensions(client):
    response = client.post("/api/manuscripts", files={"file": ("payload.exe", b"bad", "application/octet-stream")})
    assert response.status_code == 415
    assert response.json()["detail"] == "Only .docx and .pdf manuscripts are supported"
```

- [ ] **Step 2: Run the security test to confirm it fails**

Run: `pytest tests/test_security.py -q`

Expected: FAIL if the endpoint accepts arbitrary uploads.

- [ ] **Step 3: Enforce size/type limits and write operational documentation**

Reject non-DOCX/PDF extensions and oversized uploads before storage, use server-generated filenames, and never execute uploaded content. README must include environment setup, `uvicorn app.main:app --reload`, sample request flow, fallback semantic evaluator behavior, optional LLM adapter configuration, data-directory location, and the explicit limitation that OCR/copyright/visual-quality checks require later integrations or human review.

- [ ] **Step 4: Run all tests and manual smoke test**

Run: `pytest -q && uvicorn app.main:app --port 8000`

Expected: all tests PASS; opening `http://127.0.0.1:8000/` shows the upload/review workbench.

- [ ] **Step 5: Commit documentation and hardening**

```bash
git add README.md app/main.py tests/test_security.py
git commit -m "docs: describe and harden review MVP"
```

## Plan self-review

- Product coverage: template import/versioning (Task 3), DOCX/PDF and anchors (Task 4), deterministic/semantic review (Tasks 5–6), human override and dual language reports (Tasks 7–8), and security/operation (Task 9) are all covered.
- MVP scope deliberately defers true OCR, image-quality scoring, third-party reference validation, authentication, multi-journal tenancy UI, asynchronous queue infrastructure, and PDF/DOCX report export. The MVP exposes their needed data boundaries and flags them as manual review instead of fabricating results.
- The plan has no placeholder markers and all code-bearing tasks specify paths, tests, commands and implementation contracts.
