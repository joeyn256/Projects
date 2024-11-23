\documentclass[11pt, oneside]{article}    % Standard article class, use "amsart" for AMSLaTeX format
\usepackage{geometry}                      % Layout management
\geometry{letterpaper}                     % Use letter size paper
%\geometry{landscape}                      % Use landscape layout if necessary
\usepackage{graphicx}                      % Graphics support (only needed if you're including images)
\usepackage{amssymb}                       % Symbol support (only needed if you're using special symbols)
\usepackage{titlesec}                      % Custom section and title formatting
\usepackage{parskip}                       % Paragraph spacing instead of indentation

% Custom section formatting
\titleformat{\section}[hang]{\large\bfseries}{\thesection}{1em}{}
\titleformat{\subsection}[runin]{\bfseries}{\thesubsection}{1em}{}

% Title, author, and date setup
\title{\textbf{Projects}}
\author{Joseph Nihill}
\date{\today}

\begin{document}

\maketitle  % Creates title

\section*{VBA/Macro Coding Project (Fall 2024)}

I invite you to check out my \textbf{Efficient Sorting and Organization of Climbing Data using VBA Code}.

\vspace{\baselineskip}  % Adds one blank line of vertical space
\noindent
In this project, I utilized popular sorting algorithms, such as MergeSort, in Excel VBA to efficiently sort and filter climbing data. I created advanced sorting and filtering mechanisms that go beyond Excel’s built-in tables and filters, enabling more effective data presentation. Additionally, I implemented cleaning functions and applied various strategies to showcase my VBA skills.

\section*{Coding Project (Spring 2023)}

I wrote a program that parses the New York Times website and selects a random new article from their homepage. This program updates daily, ensuring a new article is chosen each day.

\vspace{\baselineskip}  % Adds one blank line of vertical space
\noindent
The second project is a helper tool that deletes unnecessary files from your personal computer. It includes multiple fail-safes to prevent accidental deletion of important files.

\end{document}