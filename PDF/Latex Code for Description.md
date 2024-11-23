\documentclass[11pt, oneside]{article}    % Standard article class, use "amsart" for AMSLaTeX format
\usepackage{geometry}                      % Layout management
\geometry{letterpaper}                     % Use letter size paper               % Use landscape layout if necessary
\usepackage{graphicx}                      % Graphics support (only needed if you're including images)
\usepackage{amssymb}                       % Symbol support (only needed if you're using special symbols)
\usepackage{titlesec}                      % Custom section and title formatting
\usepackage{parskip}                       % Paragraph spacing instead of indentation
\usepackage{enumitem}
\titleformat{\section}[hang]{\large\bfseries}{\thesection}{1em}{}
\titleformat{\subsection}[runin]{\bfseries}{\thesubsection}{1em}{}
%\title{\huge\textbf{Projects}}
\author{Joseph Nihill}
\date{\today}
\begin{document}
%\maketitle  % Creates title
\begin{center}
    {\huge \bfseries Projects}
    
    {\large \bfseries Personal VBA/Macro Coding Project (Fall 2024)}
\end{center}
\vspace{3pt}
\textbf{Efficient Sorting, Filtering, and Organiation of Climbing Data using VBA Code}
\noindent
In this project, I utilized popular sorting algorithms, such as MergeSort, in Excel VBA to efficiently sort and filter climbing data. I created advanced sorting and filtering mechanisms that go beyond Excel’s built-in tables and filters, enabling more effective data presentation. Additionally, I implemented cleaning functions and applied various strategies to showcase my VBA skills.

Below is a list of the files involved in the project:
\begin{itemize}[label=\textbullet, font=\normalfont]
    \item \textbf{Outdoor Boulder Sends Presentable File.xlsm} - Main spreadsheet for organizing and presenting climbing data, showcasing advanced sorting and filtering methods.
    \item \textbf{Sorting Project; Pivot Chart and Table Code.bas} - Contains VBA code for generating pivot charts and tables to visually represent the sorted climbing data.
    \item \textbf{Sorting Project; Formatting Code.bas} - Provides the code for advanced data formatting, enhancing the presentation of climbing data beyond Excel’s standard table formatting.
    \item \textbf{Sorting Project; Main Code.bas} - The core VBA code that implements sorting algorithms like MergeSort, responsible for efficiently sorting and filtering climbing data.
    \item \textbf{modified\_mergesort.py} - Python script for implementing a modified version of MergeSort, which is used for sorting climbing data more effectively.
    \item \textbf{mergesort.py} - Python script for implementing the basic MergeSort algorithm, providing an alternative sorting method for climbing data.
\end{itemize}
\vspace{10pt}
\begin{center}
    {\large \bfseries New York Times Coding Project(Spring 2023)}
\end{center}

\begin{itemize}[label=\textbullet, font=\normalfont]
    \item \textbf{Open file in NY Times.py} - A Python script that parses the New York Times homepage and selects a random article. The article is updated daily, ensuring a fresh selection each day.
    \item \textbf{NY\_Times\_HTML Code.txt} - Contains the HTML structure and any necessary code snippets used by the Python script to extract data from the New York Times homepage.
\end{itemize}

\begin{center}
    {\large \bfseries File Organizer Coding Project(Fall 2023)}
\end{center}

\begin{itemize}[label=\textbullet, font=\normalfont]
	\item \textbf{extension\_deleter.py} - Python script designed to clean up directories by removing unnecessary files with specified extensions.
	\item \textbf{deleter\_directories.py} - Script for removing empty or duplicate directories from your computer’s file system.
	\item \textbf{deleter\_adv\_documents.py} - A Python script that scans for and deletes redundant documents, helping to declutter your directories by removing duplicate files.
\end{itemize}

\end{document}