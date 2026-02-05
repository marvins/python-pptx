Math
====

The following classes provide access to mathematical equations and OMML (Office Math Markup Language) functionality in PowerPoint presentations.

Overview
--------

The math module provides a high-level interface for adding mathematical equations to PowerPoint slides. It handles the complex XML structure required by PowerPoint, including the `mc:AlternateContent` wrapper, `a14:m` extension elements, and automatic formatting with Cambria Math font.

Key Features
------------

* **Automatic Formatting**: OMML elements are automatically formatted with proper PowerPoint-compatible properties
* **Font Management**: Cambria Math font and language attributes applied automatically
* **Color Support**: Text properties include color scheme integration
* **Namespace Handling**: Complex PowerPoint namespace requirements handled internally
* **Compatibility**: Includes fallback content for older PowerPoint versions

Math objects
-------------

|Math| objects provide access to the mathematical content of a math shape.

.. autoclass:: pptx.math.Math()
   :members:
   :inherited-members:

   The Math class provides high-level access to OMML content within a math shape. It handles automatic formatting of text runs with PowerPoint-compatible properties including Cambria Math font, language settings, and color schemes.

   **Key Methods:**

   * ``get_omml()`` - Get the OMML XML string for the equation
   * ``set_omml(omml_xml)`` - Replace the current OMML with new XML content
   * ``add_omml(omml_xml)`` - Add OMML content with automatic formatting applied

   **Example Usage:**

   .. code-block:: python

      # Create a math equation
      math_shape = slide.shapes.add_math_equation()
      math_shape.math.add_omml('<m:oMath><m:r><m:t>x = 1</m:t></m:r></m:oMath>')

MathShape objects
------------------

|MathShape| objects represent mathematical equations on a slide.

.. autoclass:: pptx.shapes.math.MathShape()
   :members:
   :inherited-members:

   MathShape provides the interface between PowerPoint shapes and mathematical content. Each MathShape contains a Math object that manages the OMML content.

   **Key Properties:**

   * ``math`` - Access to the Math object for OMML manipulation

   **Example Usage:**

   .. code-block:: python

      # Access math content
      omml_xml = math_shape.math.get_omml()
      math_shape.math.set_omml(new_omml)

ShapeTree Integration
-------------------

.. automethod:: pptx.shapes.shapetree.ShapeTree.add_math_equation()

   Add a mathematical equation to the slide. This method creates a textbox shape with math extension markers and returns a MathShape object.

   **Parameters:**

   * ``left`` (Length, optional) - Horizontal position (default: 1 inch)
   * ``top`` (Length, optional) - Vertical position (default: 0.75 inch)
   * ``width`` (Length, optional) - Width of equation box (default: 2 inches)
   * ``height`` (Length, optional) - Height of equation box (default: 1 inch)

   **Returns:**

   * ``MathShape`` object for accessing and manipulating the equation

   **Example Usage:**

   .. code-block:: python

      # Add equation to slide
      math_shape = slide.shapes.add_math_equation(
          left=Inches(2),
          top=Inches(1),
          width=Inches(4),
          height=Inches(1)
      )

      # Add OMML content
      math_shape.math.add_omml('<m:oMath><m:r><m:t>x² + y² = z²</m:t></m:r></m:oMath>')

OMML Structure Requirements
-----------------------

PowerPoint requires a specific XML structure for mathematical equations:

**Required Structure:**

.. code-block:: xml

   <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
     <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
       <p:sp>
         <!-- PowerPoint shape with txBox="1" -->
         <p:txBody>
           <a:p>
             <a14:m>
               <m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                 <m:oMath>
                   <!-- Your equation content here -->
                 </m:oMath>
               </m:oMathPara>
             </a14:m>
           </a:p>
         </p:txBody>
       </p:sp>
     </mc:Choice>
     <mc:Fallback>
       <!-- Fallback content for older PowerPoint versions -->
     </mc:Fallback>
   </mc:AlternateContent>

**Key Components:**

* ``mc:AlternateContent`` - Compatibility wrapper for Office 2010+ features
* ``a14:m`` - Office 2010 extension that allows OMML in DrawingML
* ``m:oMathPara`` - Container for equation formatting and alignment
* ``m:oMath`` - The actual mathematical content
* ``a:rPr`` - Run properties for font, color, and language formatting

**Automatic Formatting:**

The Math class automatically applies:

* **Font**: Cambria Math with proper panose values
* **Language**: en-US locale setting
* **Color**: Integration with PowerPoint color schemes
* **Size**: Appropriate text sizing for equations

Working Examples
---------------

**Simple Equation:**

.. code-block:: python

   # Basic equation
   prs = Presentation()
   slide = prs.slides.add_slide(prs.slide_layouts[6])
   math_shape = slide.shapes.add_math_equation()
   math_shape.math.add_omml('<m:oMath><m:r><m:t>E = mc²</m:t></m:r></m:oMath>')
   prs.save('equation.pptx')

**Complex Equation with Fraction:**

.. code-block:: python

   # Fraction using OMML structure
   fraction_omml = '''
   <m:oMath>
     <m:f>
       <m:num><m:r><m:t>x + y</m:t></m:r></m:num>
       <m:den><m:r><m:t>2</m:t></m:r></m:den>
     </m:f>
   </m:oMath>
   '''

   math_shape.math.add_omml(fraction_omml)

**Quadratic Formula (Complete Example):**

.. code-block:: python

   # Complex quadratic formula with square root
   quadratic_omml = '''
   <m:oMath>
     <m:r><m:t>x = </m:t></m:r>
     <m:f>
       <m:num>
         <m:r><m:t>-b ± </m:t></m:r>
         <m:rad>
           <m:radPr/>
           <m:deg/>
           <m:e>
             <m:r><m:t>b² - 4ac</m:t></m:r>
           </m:e>
         </m:rad>
       </m:num>
       <m:den>
         <m:r><m:t>2a</m:t></m:r>
       </m:den>
     </m:f>
   </m:oMath>
   '''

   math_shape.math.add_omml(quadratic_omml)

Important Notes
--------------

* **Namespace Handling**: The Math class automatically handles complex PowerPoint namespace requirements
* **Font Requirements**: PowerPoint requires Cambria Math for proper equation rendering
* **Compatibility**: The `mc:AlternateContent` structure ensures compatibility with PowerPoint 2010+
* **Color Schemes**: Use ``a:schemeClr`` for integration with presentation themes
* **Validation**: OMML content is validated to ensure proper structure

Limitations
-----------

* **PowerPoint-Specific**: The current implementation is optimized for PowerPoint's requirements
* **XML Complexity**: Direct OMML manipulation requires understanding of PowerPoint's XML structure
* **Font Dependencies**: Cambria Math font must be available for proper rendering

See Also
--------

* :doc:`/api/shapes.html` - Shape collection and shape tree methods
* :doc:`/api/oxml.html` - Low-level XML element classes
* :doc:`/api/dml.html` - DrawingML formatting properties
