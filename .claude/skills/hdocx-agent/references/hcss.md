# H-CSS Examples

Edit `agent.edits.hcss` inside the `.hdocx` bundle.

## Named Sets

```css
@hdocx-set body {
  select: [data-hdocx-type="paragraph"];
}

@hdocx-edit mode(paragraph-formatting);

body {
  hdocx-line-spacing: 1.5;
  hdocx-first-line-indent: 2char;
}
```

## Style, List, and Part Targeting

```css
@hdocx-set body-style {
  select: style(BodyText);
}

@hdocx-set level-zero-list {
  select: list(1, 0);
}

@hdocx-set header-paragraphs {
  select: part(/word/header1.xml, paragraph);
}
```

## Run Formatting

```css
@hdocx-edit mode(run-formatting);

#r-000001 {
  hdocx-font-size: 12pt;
  hdocx-bold: true;
}
```

## Paragraph Formatting

```css
@hdocx-edit mode(paragraph-formatting);

#p-000001 {
  hdocx-align: center;
  hdocx-left-indent: 0.5in;
  hdocx-first-line-indent: 2char;
  hdocx-space-before: 6pt;
  hdocx-space-after: 6pt;
}
```

## Optional Selectors

Use `allow-empty: true` only when a missing match is acceptable:

```css
@hdocx-set optional-notes {
  select: .maybe-note;
  allow-empty: true;
}
```

Avoid broad selectors unless the user requested document-wide changes and you
have inspected the matched set.
