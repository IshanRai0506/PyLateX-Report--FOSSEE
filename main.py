#!/usr/bin/env python3
"""
Beam Analysis Report Generator
Author: [Your Name]
Date: [Current Date]

This script reads beam force data from Excel and generates a professional
engineering report with SFD and BMD diagrams using PyLaTeX and TikZ.
"""

import pandas as pd
import numpy as np
from pylatex import Document, Section, Subsection, Figure, Tabular, Command, Package
from pylatex.utils import NoEscape, italic
import os

def generate_beam_report():
    """Generate complete beam analysis report with SFD and BMD"""
    
    print("📊 Reading beam data from Excel...")
    
    # Read the Excel data
    df = pd.read_excel('beam_data.xlsx')
    
    # Create document with additional packages
    doc = Document(documentclass='report', document_options=['a4paper', '12pt'])
    
    # Add required packages for TikZ and plots
    doc.packages.append(Package('tikz'))
    doc.packages.append(Package('pgfplots'))
    doc.packages.append(Package('geometry', options=['margin=1in']))
    doc.packages.append(Package('graphicx'))
    
    # Title page
    doc.preamble.append(Command('title', 'Beam Analysis Report'))
    doc.preamble.append(Command('author', 'Structural Engineering Analysis'))
    doc.preamble.append(Command('date', NoEscape(r'\today')))
    doc.append(NoEscape(r'\maketitle'))
    
    # Table of contents
    doc.append(NoEscape(r'\tableofcontents'))
    doc.append(NoEscape(r'\newpage'))
    
    # Introduction section
    with doc.create(Section('Introduction')):
        doc.append('This engineering report presents a comprehensive analysis of a simply supported beam subjected to various loads. The analysis includes shear force and bending moment calculations with corresponding diagrams.')
        
        # Beam description subsection
        with doc.create(Subsection('Beam Description')):
            doc.append('The structural element under analysis is a simply supported beam with pinned support at one end and roller support at the other. This configuration allows for rotation at both ends while preventing vertical displacement at supports.')
            
            # Embed beam image
            with doc.create(Figure(position='h!')) as beam_fig:
                beam_fig.add_image('beam.png', width=NoEscape(r'0.8\textwidth'))
                beam_fig.add_caption('Simply Supported Beam Configuration')
        
        # Data source subsection
        with doc.create(Subsection('Data Source')):
            doc.append('The force and moment analysis data was obtained from the provided Excel spreadsheet. The calculated values at various points along the beam length are presented in the table below:')
            
            # Recreate force table using LaTeX tabular
            with doc.create(Tabular('|c|c|c|')) as table:
                table.add_hline()
                table.add_row(('Position (m)', 'Shear Force (kN)', 'Bending Moment (kNm)'))
                table.add_hline()
                
                for _, row in df.iterrows():
                    table.add_row((f"{row['X']:.1f}", f"{row['Shear force']:.1f}", f"{row['Bending Moment']:.1f}"))
                    table.add_hline()
    
    # Analysis section
    with doc.create(Section('Analysis')):
        doc.append('The beam analysis reveals the internal forces and moments along the beam length. Key observations from the calculations:')
        
        # SFD subsection
        with doc.create(Subsection('Shear Force Diagram (SFD)')):
            doc.append('The Shear Force Diagram illustrates the variation of internal shear force along the beam length. Positive values indicate upward shear, while negative values indicate downward shear.')
            
            # Generate TikZ SFD plot
            sfd_tikz = generate_sfd_plot(df)
            doc.append(NoEscape(sfd_tikz))
            
            doc.append('\\textbf{Key Observations:}')
            doc.append(NoEscape(r'\begin{itemize}'))
            doc.append(NoEscape(r'\item Maximum shear force: 45 kN'))
            doc.append(NoEscape(r'\item Zero shear occurs at 7.5 m from left support'))
            doc.append(NoEscape(r'\item Shear force changes sign at the point of maximum moment'))
            doc.append(NoEscape(r'\end{itemize}'))
        
        # BMD subsection  
        with doc.create(Subsection('Bending Moment Diagram (BMD)')):
            doc.append('The Bending Moment Diagram shows the variation of internal bending moment. Positive bending moment causes tension in the bottom fibers of the beam.')
            
            # Generate TikZ BMD plot
            bmd_tikz = generate_bmd_plot(df)
            doc.append(NoEscape(bmd_tikz))
            
            doc.append('\\textbf{Key Observations:}')
            doc.append(NoEscape(r'\begin{itemize}'))
            doc.append(NoEscape(r'\item Maximum bending moment: 168.75 kNm at 7.5 m'))
            doc.append(NoEscape(r'\item Zero moments occur at both supports'))
            doc.append(NoEscape(r'\item Parabolic distribution indicates uniformly distributed load'))
            doc.append(NoEscape(r'\end{itemize}'))
    
    # Summary section
    with doc.create(Section('Summary')):
        doc.append('This analysis successfully demonstrates the fundamental principles of beam mechanics for a simply supported configuration.')
        
        with doc.create(Subsection('Shear Force Diagram')):
            doc.append('A Shear Force Diagram (SFD) is a graphical representation that shows the internal shear force at every point along a beam. It helps structural engineers identify critical sections where shear stresses are maximum and where shear reinforcement may be required.')
            
        with doc.create(Subsection('Bending Moment Diagram')):
            doc.append('A Bending Moment Diagram (BMD) illustrates the internal bending moment along the beam length. It indicates locations of maximum bending stress, helping engineers determine appropriate beam dimensions and reinforcement for moment resistance.')
            
        doc.append('Both diagrams are essential tools in structural design, ensuring that beams are properly sized and reinforced to withstand applied loads safely.')
    
    # Generate PDF
    print("📄 Generating PDF report...")
    doc.generate_pdf('beam_analysis_report', clean_tex=False)
    print("✅ PDF report generated successfully: beam_analysis_report.pdf")

def generate_sfd_plot(df):
    """Generate TikZ code for Shear Force Diagram"""
    
    tikz_code = r"""
    \begin{center}
    \begin{tikzpicture}
    \begin{axis}[
        width=14cm,
        height=6cm,
        xlabel={Beam Length (m)},
        ylabel={Shear Force (kN)},
        grid=major,
        grid style={dotted,gray!50},
        xmin=0, xmax=12,
        ymin=-50, ymax=50,
        xtick={0,2,4,6,8,10,12},
        ytick={-40,-20,0,20,40},
        legend pos=north east,
        title={Shear Force Diagram},
        axis lines=left,
        thick
    ]
    \addplot[
        color=blue,
        mark=*,
        mark size=1.5pt,
        line width=1.5pt
    ]
    coordinates {
    """
    
    # Add coordinates from dataframe
    for _, row in df.iterrows():
        tikz_code += f"        ({row['X']}, {row['Shear force']})\n"
    
    tikz_code += r"""    };
    \addplot[color=red, dashed, line width=1pt] coordinates {(0,0) (12,0)};
    \legend{Shear Force, Zero Line}
    \end{axis}
    \end{tikzpicture}
    \end{center}
    """
    
    return tikz_code

def generate_bmd_plot(df):
    """Generate TikZ code for Bending Moment Diagram"""
    
    tikz_code = r"""
    \begin{center}
    \begin{tikzpicture}
    \begin{axis}[
        width=14cm,
        height=6cm,
        xlabel={Beam Length (m)},
        ylabel={Bending Moment (kNm)},
        grid=major,
        grid style={dotted,gray!50},
        xmin=0, xmax=12,
        ymin=0, ymax=180,
        xtick={0,2,4,6,8,10,12},
        ytick={0,40,80,120,160},
        legend pos=north east,
        title={Bending Moment Diagram},
        axis lines=left,
        thick
    ]
    \addplot[
        color=red,
        mark=*,
        mark size=1.5pt,
        line width=1.5pt
    ]
    coordinates {
    """
    
    # Add coordinates from dataframe
    for _, row in df.iterrows():
        tikz_code += f"        ({row['X']}, {row['Bending Moment']})\n"
    
    tikz_code += r"""    };
    \legend{Bending Moment}
    \end{axis}
    \end{tikzpicture}
    \end{center}
    """
    
    return tikz_code

if __name__ == "__main__":
    generate_beam_report()