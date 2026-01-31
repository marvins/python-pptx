Working with Math Equations
==========================

python-pptx supports adding mathematical equations to slides using OMML (Office Math Markup Language). This allows you to include complex mathematical formulas, fractions, superscripts, radicals, and more in your PowerPoint presentations.

.. warning::
   **Important:** This OMML feature is a stopgap implementation that provides direct access to PowerPoint's native math format. Future versions may include LaTeX-to-OMML conversion for more convenient equation input. For now, you'll need to work directly with OMML XML or use external tools to convert LaTeX to OMML.

.. note::
   OMML is the native PowerPoint math format. If you need LaTeX support, you'll need to convert LaTeX to OMML first using external tools.

Adding Math Equations
---------------------

Basic math equation
~~~~~~~~~~~~~~~~~~~~

To add a simple math equation to a slide::

    from pptx import Presentation

    prs = Presentation()
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)

    # Add a math equation using OMML XML
    omml_xml = """
    <m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
      <m:r><m:t>E = mc²</m:t></m:r>
    </m:oMath>
    """

    math_shape = slide.shapes.add_math_equation()
    math_shape.math.add_omml(omml_xml)

    prs.save('math_equation.pptx')

Positioning and sizing
~~~~~~~~~~~~~~~~~~~~~

Math shapes behave like other shapes and can be positioned and sized::

    math_shape.left = 100000  # 1 inch in EMUs
    math_shape.top = 100000   # 1 inch in EMUs
    math_shape.width = 200000 # 2 inches in EMUs
    math_shape.height = 100000 # 1 inch in EMUs

Complex Equations
-----------------

Fractions
~~~~~~~~~

To create a fraction like x/2::

    omml_xml = """
    <m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
      <m:f>
        <m:num>
          <m:r><m:t>x</m:t></m:r>
        </m:num>
        <m:den>
          <m:r><m:t>2</m:t></m:r>
        </m:den>
      </m:f>
    </m:oMath>
    """

Superscripts and Subscripts
~~~~~~~~~~~~~~~~~~~~~~~~~~~

For superscripts like x²::

    omml_xml = """
    <m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
      <m:r><m:t>x</m:t></m:r>
      <m:sSup>
        <m:e>
          <m:r><m:t>2</m:t></m:r>
        </m:e>
      </m:sSup>
    </m:oMath>
    """

For subscripts like x₁::

    omml_xml = """
    <m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
      <m:r><m:t>x</m:t></m:r>
      <m:sSub>
        <m:e>
          <m:r><m:t>1</m:t></m:r>
        </m:e>
      </m:sSub>
    </m:oMath>
    """

Radicals (Square Roots)
~~~~~~~~~~~~~~~~~~~~~~~

For square roots::

    omml_xml = """
    <m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
      <m:rad>
        <m:radPr/>
        <m:deg/>
        <m:e>
          <m:r><m:t>x</m:t></m:r>
        </m:e>
      </m:rad>
    </m:oMath>
    """

For nth roots (like cube root)::

    omml_xml = """
    <m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
      <m:rad>
        <m:radPr/>
        <m:deg>
          <m:r><m:t>3</m:t></m:r>
        </m:deg>
        <m:e>
          <m:r><m:t>x</m:t></m:r>
        </m:e>
      </m:rad>
    </m:oMath>
    """

N-ary Operators (Summation, Integration)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For summation::

    omml_xml = """
    <m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
      <m:nary>
        <m:naryPr>
          <m:chr val="∑"/>
          <m:limLoc val="undOvr"/>
        </m:naryPr>
        <m:sub>
          <m:r><m:t>i=1</m:t></m:r>
        </m:sub>
        <m:sup>
          <m:r><m:t>n</m:t></m:r>
        </m:sup>
        <m:e>
          <m:r><m:t>i</m:t></m:r>
        </m:e>
      </m:nary>
    </m:oMath>
    """

For integration::

    omml_xml = """
    <m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
      <m:nary>
        <m:naryPr>
          <m:chr val="∫"/>
          <m:limLoc val="subSup"/>
        </m:naryPr>
        <m:sub>
          <m:r><m:t>0</m:t></m:r>
        </m:sub>
        <m:sup>
          <m:r><m:t>∞</m:t></m:r>
        </m:sup>
        <m:e>
          <m:r><m:t>e^{-x^2} dx</m:t></m:r>
        </m:e>
      </m:nary>
    </m:oMath>
    """

Complex Example: Quadratic Formula
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

A complete quadratic formula::

    omml_xml = """
    <m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
      <m:r><m:t>x = </m:t></m:r>
      <m:f>
        <m:num>
          <m:r><m:t>-b</m:t></m:r>
          <m:rad>
            <m:radPr/>
            <m:deg/>
            <m:e>
              <m:r><m:t>b</m:t></m:r>
              <m:sSup>
                <m:e><m:r><m:t>2</m:t></m:r></m:e>
              </m:sSup>
              <m:r><m:t> - 4ac</m:t></m:r>
            </m:e>
          </m:rad>
        </m:num>
        <m:den>
          <m:r><m:t>2a</m:t></m:r>
        </m:den>
      </m:f>
    </m:oMath>
    """

Working with Existing Equations
-------------------------------

Getting OMML XML
~~~~~~~~~~~~~~~~

To retrieve the OMML XML from an existing math shape::

    omml_xml = math_shape.math.get_omml()
    print(omml_xml)

Replacing Equation Content
~~~~~~~~~~~~~~~~~~~~~~~~~~~

To replace the content of an existing math equation::

    new_omml = """
    <m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
      <m:r><m:t>y = mx + b</m:t></m:r>
    </m:oMath>
    """
    
    math_shape.math.set_omml(new_omml)

Finding Math Shapes
~~~~~~~~~~~~~~~~~~~

To find all math shapes on a slide::

    math_shapes = [shape for shape in slide.shapes if hasattr(shape, 'math')]
    
    for math_shape in math_shapes:
        print(f"Found math equation: {math_shape.math.get_omml()}")

Multiple Equations Example
~~~~~~~~~~~~~~~~~~~~~~~~

Creating a slide with multiple related equations::

    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])

    # Add title
    title = slide.shapes.title
    title.text = "Pythagorean Theorem"

    # Add first equation: a² + b²
    eq1_xml = """
    <m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
      <m:r><m:t>a</m:t></m:r>
      <m:sSup>
        <m:e><m:r><m:t>2</m:t></m:r></m:e>
      </m:sSup>
      <m:r><m:t> + </m:t></m:r>
      <m:r><m:t>b</m:t></m:r>
      <m:sSup>
        <m:e><m:r><m:t>2</m:t></m:r></m:e>
      </m:sSup>
    </m:oMath>
    """

    math1 = slide.shapes.add_math_equation()
    math1.math.add_omml(eq1_xml)
    math1.left = Inches(1)
    math1.top = Inches(2)

    # Add second equation: = c²
    eq2_xml = """
    <m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math">
      <m:r><m:t>= </m:t></m:r>
      <m:r><m:t>c</m:t></m:r>
      <m:sSup>
        <m:e><m:r><m:t>2</m:t></m:r></m:e>
      </m:sSup>
    </m:oMath>
    """

    math2 = slide.shapes.add_math_equation()
    math2.math.add_omml(eq2_xml)
    math2.left = Inches(4)
    math2.top = Inches(2)

    prs.save('pythagorean_theorem.pptx')

OMML Reference
---------------

Common OMML Elements
~~~~~~~~~~~~~~~~~~~

Here are the most commonly used OMML elements:

- ``m:oMath`` - Root element for a math equation
- ``m:r`` - Text run (contains text)
- ``m:t`` - Text content
- ``m:f`` - Fraction
- ``m:num`` - Numerator
- ``m:den`` - Denominator
- ``m:sSup`` - Superscript
- ``m:sSub`` - Subscript
- ``m:e`` - Expression (base of superscript/subscript)
- ``m:rad`` - Radical (square root)
- ``m:radPr`` - Radical properties
- ``m:deg`` - Degree (for nth roots)
- ``m:nary`` - N-ary operator (summation, integral)
- ``m:naryPr`` - N-ary operator properties
- ``m:chr`` - Character for n-ary operators
- ``m:sub`` - Subscript for n-ary operators
- ``m:sup`` - Superscript for n-ary operators

Namespace
~~~~~~~~~~

All OMML elements must use the math namespace::

    xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math"

Tips and Best Practices
------------------------

1. **Use proper OMML structure** - Always validate your OMML XML against the Office Open XML specification
2. **Test in PowerPoint first** - Create complex equations in PowerPoint, save as .pptx, then extract the OMML for use in python-pptx
3. **Keep equations readable** - Use proper indentation in your OMML XML strings for maintainability
4. **Handle special characters** - Use proper OMML elements instead of Unicode math characters
5. **Consider file size** - Complex equations with many nested structures can increase file size

Future Enhancements
------------------

Future versions of python-pptx may include:

- **LaTeX to OMML conversion** - Automatic conversion from LaTeX syntax to OMML
- **Equation templates** - Pre-built templates for common mathematical formulas
- **MathML support** - Import/export support for MathML format
- **Visual equation builder** - Programmatic building blocks for complex equations

For now, the direct OMML approach provides reliable access to PowerPoint's full mathematical capabilities.
