---
layout: default
title: Downloads
permalink: /downloads/
---

<section class="downloads">
  <div class="container">
    <h1 class="section-title">Download the course files</h1>
    <p class="section-subtitle">
      Every module's <strong>working</strong> file (start here) and <strong>solution</strong> file,
      plus the capstone and the raw dataset. Each is a standard <code>.xlsx</code> you can open in
      Excel 2016 or newer.
    </p>

    <div class="downloads__notice">
      <strong>100% synthetic data.</strong> All names, customers, products, reps and amounts are
      generated from a fixed seed — nothing here is real or confidential. Regenerate it yourself with
      <code>scripts/generate_dataset.py</code>.
    </div>

    <table class="downloads__table">
      <thead>
        <tr><th>Stage</th><th>Working file</th><th>Solution file</th></tr>
      </thead>
      <tbody>
        {% for m in site.data.course.modules %}
        <tr>
          <td><strong>M{{ m.number }}</strong> — {{ m.title }}</td>
          <td><a download href="{{ '/files/working/module-' | append: m.number | append: '.xlsx' | relative_url }}">module-{{ m.number }}.xlsx</a></td>
          <td><a download href="{{ '/files/solutions/module-' | append: m.number | append: '.xlsx' | relative_url }}">module-{{ m.number }}.xlsx</a></td>
        </tr>
        {% endfor %}
        <tr>
          <td><strong>🎯 Capstone</strong> — fresh file</td>
          <td><a download href="{{ '/files/working/capstone.xlsx' | relative_url }}">capstone.xlsx</a></td>
          <td><a download href="{{ '/files/solutions/capstone.xlsx' | relative_url }}">capstone.xlsx</a></td>
        </tr>
      </tbody>
    </table>

    <h2 class="section-title">For trainers</h2>
    <p>
      Delivering this course? The <strong>Trainer's Guide</strong> is a facilitator handbook —
      timed run-sheet, per-module talking points &amp; pitfalls, the dataset reference, the full
      capstone answer key, and every Microsoft documentation link in one place.
    </p>
    <ul class="downloads__list">
      <li><a download href="{{ '/files/excel-for-sales-trainer-guide.pdf' | relative_url }}">excel-for-sales-trainer-guide.pdf</a> — facilitator handbook (PDF, 11 pages)</li>
    </ul>

    <h2 class="section-title">Raw dataset (synthetic)</h2>
    <ul class="downloads__list">
      <li><a download href="{{ '/files/source/sales_data.csv' | relative_url }}">sales_data.csv</a> — the order-line export (2,000 rows)</li>
      <li><a download href="{{ '/files/source/reps.csv' | relative_url }}">reps.csv</a> — Sales Rep reference (manager, annual quota)</li>
      <li><a download href="{{ '/files/source/brands.csv' | relative_url }}">brands.csv</a> — Brand reference (category, brand manager)</li>
    </ul>

    <p class="downloads__back"><a href="{{ '/' | relative_url }}">← Back to the course</a></p>
  </div>
</section>
