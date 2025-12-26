# Dependency Audit Report

**Project:** codex-test
**Date:** 2025-12-26
**Auditor:** Automated Security Analysis

---

## Executive Summary

| Category | Count | Status |
|----------|-------|--------|
| Direct Dependencies | 2 | Needs attention |
| Security Vulnerabilities | 7 | **CRITICAL** |
| Outdated Packages | 3 | Requires upgrade |
| Missing Dependencies | 1 | Must install |
| Code Quality Issues | 2 | Should fix |

---

## 1. Direct Dependencies Analysis

### Required by Project Scripts

| Package | Required | Installed | Latest | Status |
|---------|----------|-----------|--------|--------|
| openpyxl | Yes | **NOT INSTALLED** | 3.1.5 | Install required |
| python-dateutil | Yes | 2.9.0.post0 | 2.9.0.post0 | Current |

### Transitive Dependencies

| Package | Required By | Installed | Status |
|---------|-------------|-----------|--------|
| six | python-dateutil | 1.16.0 | Current (legacy) |
| et-xmlfile | openpyxl | N/A | Will be installed |

---

## 2. Security Vulnerabilities (CRITICAL)

The following packages in the Python environment have known security vulnerabilities:

### cryptography 41.0.7 (4 vulnerabilities)

| CVE | Severity | Fixed In | Description |
|-----|----------|----------|-------------|
| PYSEC-2024-225 | High | 42.0.4 | NULL pointer dereference in pkcs12.serialize_key_and_certificates |
| CVE-2023-50782 | High | 42.0.0 | RSA key exchange vulnerability allowing message decryption |
| CVE-2024-0727 | Medium | 42.0.2 | PKCS12 NULL pointer dereference causing DoS |
| GHSA-h4gh-qq45-vh27 | High | 43.0.1 | OpenSSL vulnerability in statically linked copy |

**Recommendation:** Upgrade to `cryptography>=46.0.3`

### pip 24.0 (1 vulnerability)

| CVE | Severity | Fixed In | Description |
|-----|----------|----------|-------------|
| CVE-2025-8869 | Medium | 25.3 | Tar archive extraction path traversal vulnerability |

**Recommendation:** Upgrade to `pip>=25.3`

### setuptools 68.1.2 (2 vulnerabilities)

| CVE | Severity | Fixed In | Description |
|-----|----------|----------|-------------|
| PYSEC-2025-49 | Critical | 78.1.1 | Path traversal allowing arbitrary file writes |
| CVE-2024-6345 | Critical | 70.0.0 | Remote code execution via package_index module |

**Recommendation:** Upgrade to `setuptools>=80.9.0`

---

## 3. Code Quality Issues

### Issue 1: Inline Package Installation (High)

**Location:** `create_epc_dashboard.py:1301-1306`

```python
try:
    from dateutil.relativedelta import relativedelta
except ImportError:
    import subprocess
    subprocess.check_call(['pip3', 'install', 'python-dateutil'])
    from dateutil.relativedelta import relativedelta
```

**Problem:** Installing packages at runtime is a security risk and bad practice.

**Recommendation:** Remove inline installation code and use `requirements.txt` instead.

### Issue 2: Missing Dependency File

**Problem:** No `requirements.txt`, `setup.py`, or `pyproject.toml` exists.

**Recommendation:** Use the generated `requirements.txt` file.

---

## 4. Bloat Analysis

### Unnecessary Dependencies

| Package | Reason | Impact |
|---------|--------|--------|
| six | Python 2/3 compatibility layer (Python 2 is EOL since 2020) | Low - Required by python-dateutil |
| conan | Build system not used by project | None - System package |

### Optimized Dependency Tree

The project only needs 2 direct dependencies:
```
openpyxl>=3.1.5
  └── et-xmlfile
python-dateutil>=2.9.0
  └── six>=1.5
```

**Total packages needed:** 4 (minimal footprint)

---

## 5. Recommended Actions

### Immediate (Security Critical)

```bash
# Upgrade vulnerable system packages
pip install --upgrade pip>=25.3
pip install --upgrade setuptools>=80.9.0
pip install --upgrade cryptography>=46.0.3
```

### Required for Project

```bash
# Install project dependencies
pip install -r requirements.txt
```

### Code Changes

1. **Remove inline pip install** from `create_epc_dashboard.py`:
   - Delete lines 1301-1306
   - Rely on `requirements.txt` for dependencies

2. **Add import validation** at script start:
```python
try:
    from dateutil.relativedelta import relativedelta
    from openpyxl import Workbook
except ImportError as e:
    print(f"Missing dependency: {e}")
    print("Run: pip install -r requirements.txt")
    exit(1)
```

---

## 6. Version Pinning Strategy

### Recommended `requirements.txt`:

```
openpyxl>=3.1.5,<4.0
python-dateutil>=2.9.0,<3.0
```

**Why these constraints:**
- `>=X.Y.Z` ensures security patches are included
- `<(X+1).0` prevents major version breaks

---

## 7. Continuous Monitoring

### Recommended Tools

1. **pip-audit** - Installed and used for this audit
   ```bash
   pip install pip-audit
   pip-audit
   ```

2. **safety** - Alternative vulnerability scanner
   ```bash
   pip install safety
   safety check
   ```

3. **dependabot** - For GitHub repositories (automated PRs)

---

## Appendix: Full Package Inventory

### Environment Packages (Filtered for relevance)

| Package | Version | Security Status |
|---------|---------|-----------------|
| cryptography | 41.0.7 | VULNERABLE |
| pip | 24.0 | VULNERABLE |
| setuptools | 68.1.2 | VULNERABLE |
| python-dateutil | 2.9.0.post0 | OK |
| six | 1.16.0 | OK |
| requests | 2.32.5 | OK |
| urllib3 | 2.6.1 | OK |
| certifi | 2025.11.12 | OK |

---

*Report generated by automated dependency audit*
