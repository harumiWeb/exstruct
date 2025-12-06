# 📄 **Corporate Usage License Guide: exstruct**

_Version 0.2 / Last update: {{ auto-date }}_

---

## TOC

<!-- TOC -->

- [📄 Corporate Usage License Guide: exstruct](#-corporate-usage-license-guide-exstruct)
  - [TOC](#toc)
  - [Purpose of This Document](#purpose-of-this-document)
  - [License of exstruct](#license-of-exstruct)
  - [What You ARE Allowed To Do DO](#what-you-are-allowed-to-do-do)
    - [Use in commercial or internal software](#use-in-commercial-or-internal-software)
    - [Modify the library](#modify-the-library)
    - [Redistribute inside the company](#redistribute-inside-the-company)
    - [Embed into proprietary or closed-source systems](#embed-into-proprietary-or-closed-source-systems)
    - [In simple terms:](#in-simple-terms)
  - [Required Compliance Points MUST](#required-compliance-points-must)
    - [1 Preserve copyright notices](#1-preserve-copyright-notices)
    - [2 Do not remove or alter the license text](#2-do-not-remove-or-alter-the-license-text)
    - [3 Do not use the author's name for endorsement](#3-do-not-use-the-authors-name-for-endorsement)
  - [What You MUST NOT Do DON’T](#what-you-must-not-do-dont)
    - [Use the author's name for advertising](#use-the-authors-name-for-advertising)
    - [Distribute the library without the license](#distribute-the-library-without-the-license)
  - [License Compatibility of Dependencies](#license-compatibility-of-dependencies)
  - [Redistribution Inside the Company](#redistribution-inside-the-company)
  - [Conclusion: exstruct Is Highly Suitable for Corporate Use](#conclusion-exstruct-is-highly-suitable-for-corporate-use)
  - [Appendix: Full BSD-3-Clause License Text](#appendix-full-bsd-3-clause-license-text)
  - [About This Document](#about-this-document)

<!-- /TOC -->

---

## Purpose of This Document

This document is intended to help **corporate users**—including legal teams, IT departments, compliance officers, and internal developers—understand the licensing implications of using **exstruct** within:

- Internal business systems
- Commercial products
- Proprietary applications
- Company-wide automation or RAG/AI pipelines

It summarizes what is permitted, what is required, and what to be careful about when using exstruct in a corporate environment.

---

## 1. License of exstruct

**exstruct is released under the BSD-3-Clause License (Modified BSD License).**

The BSD-3-Clause license is a **very permissive open-source license**, allowing:

- commercial use
- modification
- redistribution
- integration into proprietary systems
- closed-source usage

There are **very few restrictions**, making it one of the most corporate-friendly licenses available.

---

## 2. What You ARE Allowed To Do (DO)

Thanks to the permissive nature of BSD-3-Clause, the following actions are explicitly allowed.

### Use in commercial or internal software

You may freely integrate exstruct into internal tools, production systems, or customer-facing products.

### Modify the library

You may adapt or change the code for internal purposes without releasing your modifications.

### Redistribute inside the company

Sharing the library with other teams, packaging it internally, or deploying it company-wide is allowed.

### Embed into proprietary or closed-source systems

No obligation to open your own code.

### In simple terms:

**You have maximum freedom to use exstruct within your organization.**

---

## 3. Required Compliance Points (MUST)

Although permissive, BSD-3-Clause includes three conditions you _must_ follow.

### (1) Preserve copyright notices

Include the original `LICENSE` file when redistributing the library.

### (2) Do not remove or alter the license text

The BSD license terms must remain intact in redistributed source or binary distributions.

### (3) Do not use the author's name for endorsement

You may not claim that the author endorses your product unless you obtain explicit permission.

---

## 4. What You MUST NOT Do (DON’T)

BSD-3-Clause prohibits only a small number of actions.

### Use the author's name for advertising

Example of a prohibited statement:

> “This system is officially approved by harumiWeb!”

### Distribute the library without the license

The LICENSE file must remain included.

---

## 5. License Compatibility of Dependencies

All dependencies of exstruct use **permissive licenses**, fully compatible with BSD-3-Clause.

| Dependency | License           | Compatibility                             |
| ---------- | ----------------- | ----------------------------------------- |
| numpy      | BSD-3             | ◎ Fully compatible                        |
| pandas     | BSD-3             | ◎ Fully compatible                        |
| scipy      | BSD-3             | ◎ Fully compatible                        |
| xlwings    | BSD-3             | ◎ Fully compatible                        |
| openpyxl   | MIT               | ◎ Fully compatible                        |
| pydantic   | MIT               | ◎ Fully compatible                        |
| Pillow     | HPND / permissive | ◎ Fully compatible                        |
| pypdfium2  | Apache-2.0        | ◎ Compatible; no distribution obligations |
| dev tools  | MIT / Apache      | ◎ No issues                               |

➡ **No GPL, LGPL, MPL, or other copyleft licenses are present.**

This means there is **no risk of forced open-sourcing** or contamination of proprietary software.

---

## 6. Redistribution Inside the Company

When redistributing exstruct (or modified versions) inside your organization:

| Scenario                                       | Required Action                                     |
| ---------------------------------------------- | --------------------------------------------------- |
| Distributing exstruct unchanged                | Include `LICENSE`                                   |
| Distributing a modified version                | Include `LICENSE` (recommended: note modifications) |
| Packaging into executables (PyInstaller, etc.) | Include `LICENSE` somewhere in the distribution     |

Note: You do **not** need to include licenses of dependencies unless you redistribute their source code.

---

## 7. Conclusion: exstruct Is Highly Suitable for Corporate Use

Based on the license structure:

- exstruct is **safe for commercial use**
- All dependencies are fully **permissive**
- There is **zero copyleft risk**
- Internal redistribution is allowed
- Integration into proprietary systems is allowed

**exstruct is a corporate-friendly open-source component with very low compliance risk.**

---

## Appendix: Full BSD-3-Clause License Text

```text
BSD 3-Clause License

Copyright (c) 2025, ExStruct Contributors
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its
   contributors may be used to endorse or promote products derived from
   this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
```

---

## About This Document

This guide is provided to assist legal, compliance, and IT governance teams in evaluating the use of exstruct within corporate environments.

It is not a substitute for professional legal advice; consult your internal legal team if needed.
