%%
%% jnuthesis.cls — Jinan University (暨南大学) Undergraduate Thesis
%% Version 2.1, 2026/06/14
%%
%% Faithful LaTeX port of the official JNU Word templates, verified page-for-page
%% against the Word renders:
%%   官方封面  本科毕业论文封面.docx              (cover page)
%%   官方正文  暨南大学本科毕业论文模板20260413.dot (body, 14 pages)
%%
%% Ground truth measured from Microsoft Word:
%%   margins  top/bottom 2.54cm, left 3cm, right 2cm
%%   spacing  body lines advance 32.5bp (Word "1.5 lines" snapped to the grid;
%%            font-size independent — 14bp body and 15bp 诚信 both 32.5bp)
%%   header   1.5cm from edge, 宋体 9pt + rule; footer 1.75cm, Times 9pt
%%   fonts    body 宋体 14; headings 宋体加粗 15/14; 诚信声明 楷体 15;
%%            cover main title 黑体 36
%%

\NeedsTeXFormat{LaTeX2e}
\ProvidesClass{jnuthesis}[2026/06/14 v2.1 JNU Undergraduate Thesis]

%% ============================================================================
%%  Options
%% ============================================================================

% --- fontset: store choice, apply after LoadClass ---
\def\jnu@fontset@load{\jnu@fontset@auto}

% Word declares: SimSun (宋体), SimHei (黑体), KaiTi_GB2312 (楷体_GB2312),
% Times New Roman.  "加粗" everywhere is faux-bold SimSun => AutoFakeBold.
%
% Probe macOS fonts FIRST: on a Mac, \IfFontExistsTF{SimSun} would fail through
% to a (slow, noisy) Metafont lookup once xeCJK is active.  Checking Songti SC
% first avoids that probe entirely here, while Windows/Linux still fall through.
\newcommand{\jnu@fontset@auto}{%
  \IfFontExistsTF{Songti SC}{\jnu@fontset@mac}{%
    \IfFontExistsTF{SimSun}{\jnu@fontset@windows}{%
      \IfFontExistsTF{Noto Serif CJK SC}{\jnu@fontset@noto}{\jnu@fontset@fandol}%
    }%
  }%
}

\newcommand{\jnu@fontset@windows}{%
  \setCJKmainfont{SimSun}[AutoFakeBold=2.5]%
  \setCJKsansfont{SimHei}%
  \setCJKmonofont{FangSong}%
  \setCJKfamilyfont{zhsong}{SimSun}[AutoFakeBold=2.5]%
  \setCJKfamilyfont{zhhei}{SimHei}[AutoFakeBold=2.5]%
  \IfFontExistsTF{KaiTi_GB2312}%
    {\setCJKfamilyfont{zhkai}{KaiTi_GB2312}[AutoFakeBold=2.5]}%
    {\setCJKfamilyfont{zhkai}{KaiTi}[AutoFakeBold=2.5]}%
  \setCJKfamilyfont{zhfs}{FangSong}%
}

\newcommand{\jnu@fontset@mac}{%
  \setCJKmainfont{Songti SC}[AutoFakeBold=2.5]%
  \setCJKsansfont{Heiti SC}%
  \setCJKmonofont{STFangsong}%
  \setCJKfamilyfont{zhsong}{Songti SC}[AutoFakeBold=2.5]%
  \setCJKfamilyfont{zhhei}{Heiti SC}[AutoFakeBold=2.5]%
  \setCJKfamilyfont{zhkai}{Kaiti SC}[AutoFakeBold=2.5]%
  \setCJKfamilyfont{zhfs}{STFangsong}%
}

\newcommand{\jnu@fontset@noto}{%
  \setCJKmainfont{Noto Serif CJK SC}[AutoFakeBold=2.5]%
  \setCJKsansfont{Noto Sans CJK SC}%
  \setCJKmonofont{Noto Sans CJK SC}%
  \setCJKfamilyfont{zhsong}{Noto Serif CJK SC}[AutoFakeBold=2.5]%
  \setCJKfamilyfont{zhhei}{Noto Sans CJK SC}[AutoFakeBold=2.5]%
  \setCJKfamilyfont{zhkai}{Noto Serif CJK SC}[AutoFakeBold=2.5]%
  \setCJKfamilyfont{zhfs}{Noto Serif CJK SC}%
}

\newcommand{\jnu@fontset@fandol}{%
  \setCJKmainfont{FandolSong-Regular.otf}[AutoFakeBold=2.5]%
  \setCJKsansfont{FandolHei-Regular.otf}%
  \setCJKmonofont{FandolFang-Regular.otf}%
  \setCJKfamilyfont{zhsong}{FandolSong-Regular.otf}[AutoFakeBold=2.5]%
  \setCJKfamilyfont{zhhei}{FandolHei-Regular.otf}[AutoFakeBold=2.5]%
  \setCJKfamilyfont{zhkai}{FandolKai-Regular.otf}[AutoFakeBold=2.5]%
  \setCJKfamilyfont{zhfs}{FandolFang-Regular.otf}%
}

% Bundled fonts in fonts/ (Overleaf / cross-platform final builds)
\newcommand{\jnu@fontset@bundled}{%
  \setCJKmainfont{simsun.ttc}[AutoFakeBold=2.5,Path=fonts/]%
  \setCJKsansfont{simhei.ttf}[Path=fonts/]%
  \setCJKmonofont{simfang.ttf}[Path=fonts/]%
  \setCJKfamilyfont{zhsong}{simsun.ttc}[AutoFakeBold=2.5,Path=fonts/]%
  \setCJKfamilyfont{zhhei}{simhei.ttf}[AutoFakeBold=2.5,Path=fonts/]%
  \setCJKfamilyfont{zhkai}{simkai.ttf}[AutoFakeBold=2.5,Path=fonts/]%
  \setCJKfamilyfont{zhfs}{simfang.ttf}[Path=fonts/]%
}

\DeclareOption{fontset=auto}{\def\jnu@fontset@load{\jnu@fontset@auto}}
\DeclareOption{fontset=windows}{\def\jnu@fontset@load{\jnu@fontset@windows}}
\DeclareOption{fontset=mac}{\def\jnu@fontset@load{\jnu@fontset@mac}}
\DeclareOption{fontset=noto}{\def\jnu@fontset@load{\jnu@fontset@noto}}
\DeclareOption{fontset=ubuntu}{\def\jnu@fontset@load{\jnu@fontset@noto}}
\DeclareOption{fontset=fandol}{\def\jnu@fontset@load{\jnu@fontset@fandol}}
\DeclareOption{fontset=bundled}{\def\jnu@fontset@load{\jnu@fontset@bundled}}

\DeclareOption*{\PassOptionsToClass{\CurrentOption}{ctexbook}}
\ProcessOptions\relax

%% ============================================================================
%%  Base class
%% ============================================================================

\LoadClass[scheme=chinese,fontset=none,zihao=false,openany,heading=true]{ctexbook}

% Word measures in big points (bp).  Body: 四号 14bp.  Empirically the Word
% render places every body line on a 32.5bp grid (1.5-line spacing snapped to
% the document grid — measured directly from the official PDF, font-size
% independent: 14bp body and 15bp 诚信 both advance 32.5bp per line).
\linespread{1.0}                 % neutralise ctex's default 1.3 multiplier
\renewcommand{\normalsize}{\@setfontsize\normalsize{14bp}{32.5bp}}
\normalsize

%% ============================================================================
%%  Packages
%% ============================================================================

\RequirePackage{geometry}
% Word: top/bottom 2.5cm, left 3cm, right 2cm; header 1.5cm / footer 1.75cm
% from the paper edge.
\geometry{
  a4paper,
  top=2.5cm,
  bottom=2.5cm,
  left=3cm,
  right=2cm,
  headheight=14pt,
  headsep=\dimexpr1cm-14pt\relax,  % header text sits 1.5cm from paper edge
  footskip=0.75cm,                 % footer text sits 1.75cm from paper edge
}

\RequirePackage{setspace}
\RequirePackage{fancyhdr}
\RequirePackage{graphicx}
\RequirePackage{caption}
\RequirePackage{array}
\RequirePackage{multirow}
\RequirePackage{makecell}
\RequirePackage{booktabs}
\RequirePackage{amsmath,amssymb,bm}
\RequirePackage{mathtools}
\RequirePackage{enumitem}
\RequirePackage{titletoc}
\RequirePackage{tikz}
\RequirePackage{CJKfntef}
\RequirePackage{etoolbox}
\RequirePackage{ulem}\normalem

\RequirePackage{hyperref}
\hypersetup{
  hidelinks,
  bookmarksnumbered=true,
  bookmarksopen=true,
  pdfstartview=FitH,
  unicode=true,
  breaklinks=true,
}

\RequirePackage[
  backend=biber,
  style=gb7714-2015,
  sorting=none,
  gbpub=false,
  gbnamefmt=lowercase,
]{biblatex}

%% ============================================================================
%%  Fonts
%% ============================================================================

\jnu@fontset@load
\IfFontExistsTF{Times New Roman}%
  {\setmainfont{Times New Roman}}%
  {\setmainfont{TeX Gyre Termes}}

\newcommand{\songti}{\CJKfamily{zhsong}}
\newcommand{\heiti}{\CJKfamily{zhhei}}
\newcommand{\kaishu}{\CJKfamily{zhkai}}
\newcommand{\fangsong}{\CJKfamily{zhfs}}

% Enclosed alphanumerics (①②③ … U+2460–U+24FF) live in the CJK fonts, not in
% Times New Roman; route them through the CJK font so they render (e.g. 布氏硬度①).
\xeCJKDeclareCharClass{CJK}{"2460 -> "24FF}

%% ============================================================================
%%  Word line grid: 19.05pt pitch, body snaps to 2 lines = 38.1pt
%% ============================================================================

% Body text: SimSun 14bp (四号) on the 32.5bp grid pitch.
\newcommand{\jnu@body}{\fontsize{14bp}{32.5bp}\selectfont}
% 诚信声明 body: KaiTi 15bp (小三) on the same 32.5bp grid pitch.
\newcommand{\jnu@stmtbody}{\fontsize{15bp}{32.5bp}\selectfont}
% Page titles (论文题目/Abstract title/目录): 18bp (小二).
\newcommand{\jnu@pagetitle}{\fontsize{18bp}{32.5bp}\selectfont}
% 诚信声明 title: 22bp (二号), Word line spacing fixed at 40pt.
\newcommand{\jnu@stmttitle}{\fontsize{22bp}{40bp}\selectfont}
% Chapter: 15bp (小三). Sections: 14bp (四号).
\newcommand{\jnu@chapsize}{\fontsize{15bp}{32.5bp}\selectfont}
\newcommand{\jnu@secsize}{\fontsize{14bp}{32.5bp}\selectfont}
% Captions: 10.5bp (五号).
\newcommand{\jnu@captionsize}{\fontsize{10.5bp}{15bp}\selectfont}
% Header 9bp (小五).
\newcommand{\jnu@headersize}{\fontsize{9bp}{10.8bp}\selectfont}
% Cover main title: SimHei 36bp (小初).
\newcommand{\jnu@coversize}{\fontsize{36bp}{43bp}\selectfont}
% Cover fields: 14bp.
\newcommand{\jnu@fieldsize}{\fontsize{14bp}{24bp}\selectfont}

\AtBeginDocument{\jnu@body\songti}

% Displayed equations sit on the airy Word grid; keep moderate skips.
\AtBeginDocument{%
  \setlength{\abovedisplayskip}{9pt plus 2pt minus 2pt}%
  \setlength{\belowdisplayskip}{9pt plus 2pt minus 2pt}%
  \setlength{\abovedisplayshortskip}{6pt plus 1pt minus 1pt}%
  \setlength{\belowdisplayshortskip}{6pt plus 1pt minus 1pt}%
}

%% ============================================================================
%%  Paragraph layout
%% ============================================================================

\setlength{\parindent}{2\ccwd}
\setlength{\parskip}{0pt}

\setlist{nosep}
\setlist[enumerate,1]{label={（\arabic*）},left=2\ccwd,labelsep=0.25\ccwd}
\setlist[itemize]{left=2\ccwd,labelsep=0.5\ccwd}

%% ============================================================================
%%  Page styles
%% ============================================================================

% Word puts a thin border line under the header text in both sections.
\renewcommand{\headrulewidth}{0.75pt}

% Front matter: 暨南大学本科毕业论文（设计）, no footer.
\fancypagestyle{jnu-front}{%
  \fancyhf{}%
  \fancyhead[C]{\jnu@headersize\songti 暨南大学本科毕业论文（设计）}%
  \renewcommand{\headrulewidth}{0.75pt}%
  \renewcommand{\footrulewidth}{0pt}%
}

% Main matter: running title + TNR 9pt centered page number.
\fancypagestyle{jnu-main}{%
  \fancyhf{}%
  \fancyhead[C]{\jnu@headersize\songti \@runhead}%
  \fancyfoot[C]{\fontsize{9bp}{10.8bp}\selectfont\thepage}%
  \renewcommand{\headrulewidth}{0.75pt}%
  \renewcommand{\footrulewidth}{0pt}%
}

% Cover: completely empty.
\fancypagestyle{jnu-empty}{%
  \fancyhf{}%
  \renewcommand{\headrulewidth}{0pt}%
  \renewcommand{\footrulewidth}{0pt}%
}

\pagestyle{jnu-front}

%% ============================================================================
%%  Headings (Word: SimSun bold; chapter 15pt, sections 14pt, left aligned)
%% ============================================================================

\setcounter{secnumdepth}{3}

\ctexset{
  chapter = {
    name = {,},
    number = \arabic{chapter},
    format = \jnu@chapsize\songti\bfseries\raggedright,
    aftername = \enspace,
    beforeskip = 6pt,
    afterskip = 6pt,
    fixskip = true,
    pagestyle = jnu-main,
  },
  section = {
    name = {,},
    number = \arabic{chapter}.\arabic{section},
    format = \jnu@secsize\songti\bfseries\raggedright,
    aftername = \enspace,
    beforeskip = 6pt,
    afterskip = 6pt,
    indent = 0pt,
  },
  subsection = {
    name = {,},
    number = \arabic{chapter}.\arabic{section}.\arabic{subsection},
    format = \jnu@secsize\songti\bfseries\raggedright,
    aftername = \enspace,
    beforeskip = 6pt,
    afterskip = 6pt,
    indent = 0pt,
  },
  subsubsection = {
    name = {,},
    number = \arabic{chapter}.\arabic{section}.\arabic{subsection}.\arabic{subsubsection},
    format = \jnu@secsize\songti\bfseries\raggedright,
    aftername = \enspace,
    beforeskip = 6pt,
    afterskip = 6pt,
    indent = 0pt,
  },
}

% Special unnumbered chapters.  Word: 参考文献 & 致谢 are CENTERED,
% 附录 is LEFT aligned; all SimSun 15pt bold.
\newcommand{\jnucenterchapter}[1]{%
  \clearpage
  \phantomsection
  \addcontentsline{toc}{chapter}{#1}%
  {\parindent=0pt\centering\jnu@chapsize\songti\bfseries #1\par}%
  \vspace{6pt}%
}
\newcommand{\jnuleftchapter}[1]{%
  \clearpage
  \phantomsection
  \addcontentsline{toc}{chapter}{#1}%
  {\parindent=0pt\raggedright\jnu@chapsize\songti\bfseries #1\par}%
  \vspace{6pt}%
}

%% ============================================================================
%%  Table of Contents (目 录)
%% ============================================================================

\setcounter{tocdepth}{3}

% 目录条目字号按官方 .dot 实测：L1(章) 四号14bp 加黑；L2(节) 12bp；L3 10.5bp。
% \contentspush sizes the label box automatically (no 附录A overlap).
\titlecontents{chapter}
  [0pt]
  {\addvspace{2pt}\fontsize{14bp}{24bp}\selectfont\songti\bfseries}
  {\contentspush{\thecontentslabel\enspace}}
  {}
  {\titlerule*[4pt]{.}\contentspage}

\titlecontents{section}
  [10.5bp]
  {\fontsize{12bp}{19bp}\selectfont\songti}
  {\contentspush{\thecontentslabel\enspace}}
  {}
  {\titlerule*[4pt]{.}\contentspage}

\titlecontents{subsection}
  [21bp]
  {\fontsize{10.5bp}{16.5bp}\selectfont\songti}
  {\contentspush{\thecontentslabel\enspace}}
  {}
  {\titlerule*[4pt]{.}\contentspage}

% 目录 title: 18pt bold, centered, 目 and 录 separated widely.
% The official Word template prints explanatory paragraphs between the
% title and the entries; put them in \jnutocpreamble (may be empty).
\newcommand{\jnutocpreamble}{}
\renewcommand{\tableofcontents}{%
  \clearpage
  \thispagestyle{jnu-front}%
  \begin{center}
    {\jnu@pagetitle\songti\bfseries 目\hspace{2\ccwd}录}
  \end{center}
  \vspace{6pt}%
  {\jnu@body\songti\setlength{\parindent}{2\ccwd}\jnutocpreamble\par}%
  \@starttoc{toc}%
  \clearpage
}

%% ============================================================================
%%  Captions / floats / equations
%% ============================================================================

\DeclareCaptionFont{jnucaption}{\jnu@captionsize\songti}
\captionsetup{
  font=jnucaption,
  labelfont=jnucaption,
  labelsep=space,
  justification=centering,
  format=plain,
  skip=6pt,
}

\renewcommand{\thefigure}{\arabic{chapter}-\arabic{figure}}
\renewcommand{\thetable}{\arabic{chapter}-\arabic{table}}
% Word tags displayed equations as （式 3-1） at the right margin.
\renewcommand{\theequation}{\arabic{chapter}-\arabic{equation}}
\newtagform{jnu}{（式\,}{）}
\usetagform{jnu}

\ctexset{
  figurename = {图},
  tablename = {表},
}

% Word-grid table cells: 14bp on tall lines.
\newcommand{\jnutablecell}{\fontsize{14bp}{30bp}\selectfont\songti}

%% ============================================================================
%%  Metadata
%% ============================================================================

% \biaoti sets both the 摘要 page title (\@biaoti) and the running header
% (\@runhead) to the real thesis title.  The defaults below reproduce the
% official .dot/.docx placeholder strings exactly (brackets included), so an
% unfilled template compiles to a page-for-page match of the Word original.
\newcommand{\biaoti}[1]{\gdef\@biaoti{#1}\gdef\@runhead{#1}}
\newcommand{\entitle}[1]{\gdef\@entitle{#1}}
\newcommand{\xueyuan}[1]{\gdef\@xueyuan{#1}}
\newcommand{\xuexi}[1]{\gdef\@xuexi{#1}}
\newcommand{\zhuanye}[1]{\gdef\@zhuanye{#1}}
\newcommand{\xingming}[1]{\gdef\@xingming{#1}}
\newcommand{\xuehao}[1]{\gdef\@xuehao{#1}}
\newcommand{\daoshi}[1]{\gdef\@daoshi{#1}}
\newcommand{\thesisyear}[1]{\gdef\@thesisyear{#1}}
\newcommand{\thesismonth}[1]{\gdef\@thesismonth{#1}}
\newcommand{\thesisday}[1]{\gdef\@thesisday{#1}}
\newcommand{\thesisdate}[3]{\thesisyear{#1}\thesismonth{#2}\thesisday{#3}}

\gdef\@biaoti{[在此处键入论文题目]}        % 摘要页标题占位符（与官方 .dot 一致）
\gdef\@runhead{[毕业论文（设计）题目]}      % 正文页眉占位符（与官方 .dot 一致）
\gdef\@entitle{[在此处键入论文英文题目]}    % 英文摘要页标题占位符
\xueyuan{}\xuexi{}\zhuanye{}\xingming{}\xuehao{}\daoshi{}
\thesisyear{}\thesismonth{}\thesisday{}

%% ============================================================================
%%  Cover (本科毕业论文封面.docx, one page, no header/footer)
%% ============================================================================

% A field row: bold 14pt label spread over a fixed 6-char box (so every
% underline starts at the same x), then an underline to the right margin.
% #1 = label (2 or 4 chars), #2 = content
\newcommand{\jnu@fieldrow}[2]{%
  \par\noindent
  \makebox[6\ccwd][s]{\jnu@fieldsize\songti\bfseries #1}%
  \hspace{0.4\ccwd}%
  {\jnu@fieldsize\songti\uline{\makebox[\dimexpr\linewidth-6.4\ccwd\relax][c]{#2\strut}}}%
  \par
}

\newcommand{\jnu@rendercover}{%
  \clearpage
  \thispagestyle{jnu-empty}%
  \begingroup
    \setlength{\parindent}{0pt}%
    \setstretch{1.0}%            % cover rows are single-spaced (not 1.5)
    \centering
    \vspace*{0mm}
    % 1. seal  (Word: 41.4mm wide)
    \IfFileExists{figs/jnu-logo.png}%
      {\includegraphics[width=41.4mm]{figs/jnu-logo.png}}{\rule{0pt}{43mm}}%
    \par\vspace{5mm}
    % 2. calligraphy name  (Word: 131.9mm wide)
    \IfFileExists{figs/jnu-name.png}%
      {\includegraphics[width=131.9mm]{figs/jnu-name.png}}{}%
    \par\vspace{2mm}
    % 3. main title  (SimHei 小初/36pt)
    {\jnu@coversize\heiti 本科生毕业论文\par}
    \vspace{9mm}
    % 4. thesis title row
    \jnu@fieldrow{论文题目}{\jnu@cover@biaoti}%
    \vspace{5mm}
    % 5. field rows (evenly spaced)
    \jnu@fieldrow{学院}{\jnu@cover@xueyuan}\vspace{4mm}
    \jnu@fieldrow{学系}{\jnu@cover@xuexi}\vspace{4mm}
    \jnu@fieldrow{专业}{\jnu@cover@zhuanye}\vspace{4mm}
    \jnu@fieldrow{姓名}{\jnu@cover@xingming}\vspace{4mm}
    \jnu@fieldrow{学号}{\jnu@cover@xuehao}\vspace{4mm}
    \jnu@fieldrow{指导教师}{\jnu@cover@daoshi}%
    \vspace{6mm}
    % 6. date  年 月 日
    {\jnu@fieldsize\songti\bfseries
      \makebox[2\ccwd][c]{\jnu@cover@year}年%
      \hspace{2\ccwd}\makebox[1.6\ccwd][c]{\jnu@cover@month}月%
      \hspace{2\ccwd}\makebox[1.6\ccwd][c]{\jnu@cover@day}日\par}
  \endgroup
  \clearpage
}

\newcommand{\makecover}{%
  \begingroup
    \let\jnu@cover@biaoti\@biaoti
    \let\jnu@cover@xueyuan\@xueyuan
    \let\jnu@cover@xuexi\@xuexi
    \let\jnu@cover@zhuanye\@zhuanye
    \let\jnu@cover@xingming\@xingming
    \let\jnu@cover@xuehao\@xuehao
    \let\jnu@cover@daoshi\@daoshi
    \let\jnu@cover@year\@thesisyear
    \let\jnu@cover@month\@thesismonth
    \let\jnu@cover@day\@thesisday
    \jnu@rendercover
  \endgroup
}

\newcommand{\makeblankcover}{%
  \begingroup
    \def\jnu@cover@biaoti{}\def\jnu@cover@xueyuan{}\def\jnu@cover@xuexi{}%
    \def\jnu@cover@zhuanye{}\def\jnu@cover@xingming{}\def\jnu@cover@xuehao{}%
    \def\jnu@cover@daoshi{}\def\jnu@cover@year{}\def\jnu@cover@month{}%
    \def\jnu@cover@day{}%
    \jnu@rendercover
  \endgroup
}

%% ============================================================================
%%  诚信声明 (KaiTi_GB2312: title 22pt bold, body 15pt, own page)
%% ============================================================================

\newcommand{\makestatement}{%
  \clearpage
  \thispagestyle{jnu-front}%
  \begingroup
    \setlength{\parindent}{2\ccwd}%
    \vspace*{14.5mm}
    \begin{center}
      {\jnu@stmttitle\kaishu\bfseries 诚\hspace{0.5\ccwd}信\hspace{0.5\ccwd}声\hspace{0.5\ccwd}明}
    \end{center}
    \vspace{14mm}
    {\jnu@stmtbody\kaishu
      \noindent\hspace{2\ccwd}我声明，所呈交的毕业论文（设计）是本人在老师指导下进行的研究工作及取得的研究成果。除文中已经注明引用的内容外，本论文（设计）不包含任何其他个人或集体已经发表或撰写过的作品成果。在论文（设计）撰写过程中，本人严格遵守学术道德和学术规范，恪守学术诚信，坚决杜绝任何形式的抄袭、剽窃、伪造、篡改、买卖、代写毕业论文等学术不端行为，本人自愿承担因违反上述承诺而产生的后果和责任。\par
    }
    \vspace{22mm}
    {\jnu@stmtbody\kaishu
      \noindent\hspace{1\ccwd}签名：\hspace{10\ccwd}日期：\hspace{1.5\ccwd}年\hspace{1\ccwd}月\hspace{1\ccwd}日\par
    }
  \endgroup
  \clearpage
}

%% ============================================================================
%%  Abstracts
%% ============================================================================

% 中文摘要页: title 18pt bold centered; [摘 要] run-in bold; body 14pt.
\newenvironment{zhabstract}{%
  \clearpage
  \thispagestyle{jnu-front}%
  \begin{center}
    {\jnu@pagetitle\songti\bfseries \@biaoti}
  \end{center}
  \vspace{12pt}
  \jnu@body\songti
  \setlength{\parindent}{2\ccwd}%
  \noindent\hspace{2\ccwd}{\bfseries [摘\hspace{0.5\ccwd}要]}\,%
}{%
  \par\clearpage
}

\newcommand{\zhaiyao}[1]{%
  \par\vspace{\baselineskip}%
  \noindent\hspace{2\ccwd}{\songti\bfseries [关键词]}\hspace{0.5\ccwd}#1\par
}

% 英文摘要页: Times New Roman throughout.
\newenvironment{enabstract}{%
  \clearpage
  \thispagestyle{jnu-front}%
  \begin{center}
    {\jnu@pagetitle\bfseries \@entitle}
  \end{center}
  \vspace{12pt}
  \jnu@body
  \setlength{\parindent}{2\ccwd}%
  \noindent\hspace{2\ccwd}{\bfseries Abstract：}%
}{%
  \par\clearpage
}

\newcommand{\enkeywords}[1]{%
  \par\vspace{\baselineskip}%
  \noindent\hspace{2\ccwd}{\bfseries Keywords：}\hspace{0.5\ccwd}#1\par
}

%% ============================================================================
%%  Main matter switch
%% ============================================================================

\newcommand{\jnumainmatter}{%
  \clearpage
  \pagenumbering{arabic}%
  \setcounter{page}{1}%
  \pagestyle{jnu-main}%
}

%% ============================================================================
%%  Bibliography (GB/T 7714-2015)
%% ============================================================================

% 按《撰写规范》(附件2 三(六)1)：参考文献标注采用顺序编码制，标注符号为“[ ]”、
% 作为上标、在标点符号前使用，连续多篇压缩为“[5-6]”。biblatex-gb7714-2015 默认的
% \supercite 仅输出上标数字而不带方括号，此处重定义补上方括号并启用连续压缩。
\DeclareCiteCommand{\supercite}[\mkbibsuperscript]
  {\usebibmacro{cite:init}\bibopenbracket}
  {\usebibmacro{citeindex}\usebibmacro{cite:comp}}
  {}
  {\usebibmacro{cite:dump}\bibclosebracket}

\providecommand{\upcite}[1]{\supercite{#1}}
\defbibheading{jnubib}[参考文献]{%
  \jnucenterchapter{#1}%
}

%% ============================================================================
%%  Appendix (附录A, 附录B …)
%% ============================================================================

\newcommand{\jnuappendix}{%
  \setcounter{chapter}{0}%
  \setcounter{section}{0}%
  \gdef\thechapter{\Alph{chapter}}%
  \ctexset{
    chapter/name = {附录,},
    chapter/number = \Alph{chapter},
    chapter/aftername = \enspace,
    section/number = \Alph{chapter}.\arabic{section},
    subsection/number = \Alph{chapter}.\arabic{section}.\arabic{subsection},
  }%
}

%% ============================================================================
%%  Math utilities (kept from v1)
%% ============================================================================

\DeclareMathOperator{\Cov}{Cov}
\DeclareMathOperator{\Var}{Var}
\newcommand{\diff}{\mathop{}\!\mathrm{d}}
\newcommand{\e}{\mathrm{e}}
\newcommand{\im}{\mathrm{i}}
\def\T{\mathrm{T}\kern-.5pt}

%% ============================================================================
%%  Local config
%% ============================================================================

\InputIfFileExists{jnuthesis.cfg}{\typeout{Loading jnuthesis.cfg}}{}

\endinput
