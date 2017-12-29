---
title: "WordDoc Example Scripts"
excerpt: "WordDoc Quick Start - Create Microsoft Word Documents directly from PowerShell"
layout: archive
header:
  overlay_color: "#5e616c"
  overlay_image: /assets/images/unsplash.jpg
  overlay_filter: rgba(100, 100, 100, 0.5)
  cta_label: "<i class='fa fa-download'></i> Get Started"
  cta_url: "/quick-start"
  caption: ""
permalink: /scripts
---
<div class="grid__wrapper">
  {% for post in site.scripts %}
    {% include archive-single.html type="grid" %}
  {% endfor %}
</div>
