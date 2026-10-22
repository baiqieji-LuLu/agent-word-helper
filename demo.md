# Word Helper Demo

## Introduction

This document tests the **bold**, *italic*, ~~strikethrough~~, `inline code`, and [Markdown link](https://example.com) features of the renderer. It also embeds inline math such as $\frac{a}{b}$ and the Greek letters $\alpha + \beta = \gamma$ inside a normal paragraph.

## Lists

### Unordered List

- First item
  - Nested item one
  - Nested item two
- Second item

### Ordered List

1. First step
2. Second step
3. Third step

## Quote

> A blockquote should be indented and rendered with a left border, separate from the surrounding paragraphs.

## Table

| Feature | Status | Score |
| :--- | :---: | ---: |
| Parser | Ready | 95 |
| Renderer | Testing | 88 |
| Tables | Ready | 90 |

## Code Block

```python
def add(a, b):
    return a + b


if __name__ == "__main__":
    print(add(2, 3))
```

## Mathematics

### Inline Formulas

The area of a circle is $A = \pi r^2$, and the fraction $\frac{1}{2}$ appears in the same line.

### Single Display Formula

$$\int_0^1 x^2 \mathrm{d}x = \frac{1}{3}$$

### Multi-line Aligned Formula

$$
\begin{align}
(a+b)^2 &= a^2 + 2ab + b^2 \\
(a-b)^2 &= a^2 - 2ab + b^2
\end{align}
$$
