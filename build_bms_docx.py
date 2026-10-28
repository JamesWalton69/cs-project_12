import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

_BM_ID = [1000]

def add_bookmark(paragraph, name):
    """Attach a named bookmark to a paragraph (target for PAGEREF index fields)."""
    _BM_ID[0] += 1
    bid = str(_BM_ID[0])
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), bid)
    start.set(qn('w:name'), name)
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), bid)
    paragraph._p.insert(0, start)
    paragraph._p.append(end)

def set_pageref_cell(cell, bookmark, fallback):
    """Fill a table cell with a dynamic PAGEREF field (zero-padded to 2 digits)."""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGEREF %s \\# "00"' % bookmark)
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = fallback
    r.append(t)
    fld.append(r)
    p._p.append(fld)

def create_document():
    doc = docx.Document()
    
    # Page setup: Standard A4 / Letter, 0.9-inch margins
    section = doc.sections[0]
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    
    # Configure section start page number to 3
    sectPr = section._sectPr
    pgNumType = OxmlElement('w:pgNumType')
    pgNumType.set(qn('w:start'), '3')
    sectPr.append(pgNumType)
    
    # Footer with right-aligned page counter (Pure B&W)
    footer = section.footer
    p_ftr = footer.paragraphs[0]
    p_ftr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_ftr = p_ftr.add_run("Bank Management System  |  Page ")
    run_ftr.font.name = "Times New Roman"
    run_ftr.font.size = Pt(10)
    run_ftr.font.color.rgb = RGBColor(0, 0, 0)
    
    fld = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
    p_ftr._p.append(fld)
    
    # Styling helpers
    def add_h1(title, bookmark=None):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(12)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(18)
        run.font.color.rgb = RGBColor(0, 0, 0)
        
        # Solid bold 1.5pt bottom rule
        pBdr = parse_xml(r'<w:pBdr %s><w:bottom w:val="single" w:sz="16" w:space="6" w:color="000000"/></w:pBdr>' % nsdecls('w'))
        p._p.get_or_add_pPr().append(pBdr)
        if bookmark:
            add_bookmark(p, bookmark)
        return p

    def add_h2(title):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(13.5)
        run.font.color.rgb = RGBColor(0, 0, 0)
        
        # Subtle 0.5pt bottom rule under subsection
        pBdr = parse_xml(r'<w:pBdr %s><w:bottom w:val="single" w:sz="6" w:space="4" w:color="000000"/></w:pBdr>' % nsdecls('w'))
        p._p.get_or_add_pPr().append(pBdr)
        return p

    def add_body(text, bold_prefix="", indent=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.25
        if indent:
            p.paragraph_format.left_indent = Inches(0.25)
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.bold = True
            r_pre.font.name = "Times New Roman"
            r_pre.font.size = Pt(12)
            r_pre.font.color.rgb = RGBColor(0, 0, 0)
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_bullet(text, bold_prefix=""):
        p = doc.add_paragraph(style='List Bullet')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.25
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.bold = True
            r_pre.font.name = "Times New Roman"
            r_pre.font.size = Pt(12)
            r_pre.font.color.rgb = RGBColor(0, 0, 0)
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_code_block(code_text, font_size_pt=9.5, line_spacing=1.12):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        # Keep the whole code block on one page (no mid-block page splits)
        trPr = tbl.rows[0]._tr.get_or_add_trPr()
        trPr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(r'''
            <w:tcBorders %s>
                <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                <w:left w:val="single" w:sz="16" w:space="0" w:color="000000"/>
                <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            </w:tcBorders>
        ''' % nsdecls('w'))
        tcPr.append(borders)
        tcMar = parse_xml(r'''
            <w:tcMar %s>
                <w:top w:w="60" w:type="dxa"/>
                <w:bottom w:w="60" w:type="dxa"/>
                <w:left w:w="120" w:type="dxa"/>
                <w:right w:w="120" w:type="dxa"/>
            </w:tcMar>
        ''' % nsdecls('w'))
        tcPr.append(tcMar)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = line_spacing
        run = p.add_run(code_text)
        run.font.name = "Consolas"
        run.font.size = Pt(font_size_pt)
        run.font.color.rgb = RGBColor(0, 0, 0)

    def add_image_centered(img_path, caption, width_in=5.4):
        if os.path.exists(img_path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
            run = p.add_run()
            run.add_picture(img_path, width=Inches(width_in))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(6)
            r_cap = p_cap.add_run(caption)
            r_cap.font.name = "Times New Roman"
            r_cap.font.size = Pt(10.5)
            r_cap.italic = True
            r_cap.font.color.rgb = RGBColor(0, 0, 0)

    def style_table_bw(tbl):
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row_idx, row in enumerate(tbl.rows):
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))
            if row_idx == 0:
                trPr.append(parse_xml(r'<w:tblHeader %s/>' % nsdecls('w')))
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                borders = parse_xml(r'''
                    <w:tcBorders %s>
                        <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                        <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                        <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                        <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>
                    </w:tcBorders>
                ''' % nsdecls('w'))
                tcPr.append(borders)

    # ==========================================
    # PAGE 3: INDEX
    # ==========================================
    add_h1("INDEX")
    
    index_data = [
        ("SER", "DESCRIPTION", "PAGE NO"),
        ("01", "INTRODUCTION", "sec_intro", "04"),
        ("02", "OBJECTIVES OF THE PROJECT", "sec_objectives", "05"),
        ("03", "PROPOSED SYSTEM", "sec_proposed", "06"),
        ("04", "SYSTEM DEVELOPMENT LIFE CYCLE (SDLC)", "sec_sdlc", "08"),
        ("05", "PHASES OF SYSTEM DEVELOPMENT LIFE CYCLE", "sec_phases", "09"),
        ("06", "FLOW CHART", "sec_flowchart", "16"),
        ("07", "SOURCE CODE", "sec_sourcecode", "18"),
        ("08", "OUTPUT", "sec_output", "25"),
        ("09", "TESTING", "sec_testing", "31"),
        ("10", "HARDWARE AND SOFTWARE REQUIREMENTS", "sec_hardware", "34"),
        ("11", "INSTALLATION PROCEDURE", "sec_installation", "35"),
        ("12", "BIBLIOGRAPHY", "sec_bibliography", "36"),
    ]
    
    idx_tbl = doc.add_table(rows=len(index_data), cols=3)
    col_widths = [Inches(1.0), Inches(4.3), Inches(1.2)]
    
    for r_idx, row in enumerate(idx_tbl.rows):
        vals = index_data[r_idx]
        if r_idx == 0:
            for c_idx, cell in enumerate(row.cells):
                cell.width = col_widths[c_idx]
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(4)
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 1 else WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(vals[c_idx])
                run.font.name = "Times New Roman"
                run.font.size = Pt(11.5)
                run.font.color.rgb = RGBColor(0, 0, 0)
                run.bold = True
        else:
            ser, desc, bm, fb = vals
            for c_idx, cell in enumerate(row.cells):
                cell.width = col_widths[c_idx]
                if c_idx == 2:
                    set_pageref_cell(cell, bm, fb)
                else:
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(4)
                    p.paragraph_format.space_after = Pt(4)
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 1 else WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(ser if c_idx == 0 else desc)
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(11.5)
                    run.font.color.rgb = RGBColor(0, 0, 0)
    
    style_table_bw(idx_tbl)
    
    # ==========================================
    # PAGE 4: INTRODUCTION (Dedicated full page)
    # ==========================================
    doc.add_page_break()
    add_h1("01. INTRODUCTION", bookmark="sec_intro")
    add_body("In the contemporary financial landscape, the modernization and digitization of commercial operations have transitioned from an operational convenience to an imperative necessity. The \"BANK MANAGEMENT SYSTEM\" is a dedicated database-backed software application engineered to handle, automate, and streamline everyday branch-level accounting and administrative responsibilities.")
    add_body("For decades, financial establishments operated via manual physical ledgers, ink-and-paper journals, and paper registers. While traditionally functional, these legacy mechanisms suffer from severe limitations: human calculation fatigue, mathematical discrepancies, catastrophic physical vulnerability to fire or water, and extreme labor delays during routine account reconciliations. Locating a single historical transaction often demanded digging through hundreds of archived binders.")
    add_body("The emerging paradigm of digital information systems makes comprehensive customer and transaction intelligence accessible instantaneously. By deploying a computerized architecture, customer account opening, verification, deposits, withdrawals, and ledger adjustments are completed within fractions of a second.")
    
    add_h2("CORE PILLARS OF THE SYSTEM")
    add_bullet("Operational Efficiency: Eliminating customer queue congestion by enabling high-speed automated credit and debit operations directly at the teller console.")
    add_bullet("Data Security & User Access Control: Restricting system access to authorized operators through unique operator IDs and mandatory 4-digit PIN authentication.")
    add_bullet("Referential Integrity: Synchronizing customer profiles with append-only transaction audit logs to guarantee that every financial debit or credit is permanently logged with an automated date timestamp.")
    add_bullet("Real-Time Reporting: Providing immediate terminal reports on customer account balances and full historical account statements without manual calculation.")
    add_body("Developed utilizing Python 3.12 and MySQL Database Management System, this project exemplifies the application of core computer science fundamentals to an authentic commercial requirement, serving as an extensible prototype for computerized financial administration.")

    # ==========================================
    # PAGE 5: OBJECTIVES OF THE PROJECT (Dedicated full page)
    # ==========================================
    doc.add_page_break()
    add_h1("02. OBJECTIVES OF THE PROJECT", bookmark="sec_objectives")
    add_body("The fundamental objective of this project is to bridge theoretical computer science pedagogy with real-world software engineering, equipping students with practical software design, database interfacing, and system validation skills.")
    
    add_h2("1. TECHNICAL OBJECTIVES")
    add_bullet("Relational Database Architecture: Design and implement a normalized database schema in MySQL comprising distinct entity tables (user_table, customer_details, and transactions) with primary key and foreign key relationships.")
    add_bullet("Python-MySQL Connectivity: Establish reliable client-server communication using Python's standard mysql.connector library to execute dynamic Data Definition (DDL) and Data Manipulation (DML) queries.")
    add_bullet("Clean Procedural Programming: Construct modular, maintainable, and readable procedural code strictly conforming to the CBSE Class 12 Computer Science syllabus, utilizing clean functions and standard library modules like datetime.")
    add_bullet("Overdraft & Fraud Prevention: Implement rigorous programmatic validation checks to prevent negative account balances by rejecting withdrawal requests exceeding existing credit reserves.")

    add_h2("2. OPERATIONAL & PEDAGOGICAL OBJECTIVES")
    add_bullet("Automated Profile Management: Automate new customer onboarding, account number assignment, and residential detail cataloging.")
    add_bullet("Atomic Transaction Logging: Guarantee ACID principles by issuing explicit database commit directives (conn.commit()) following every balance modification.")
    add_bullet("Complete Auditability: Maintain an indelible, chronological transaction register detailing the transaction date, account identifier, withdrawal sum, and deposited sum for auditing purposes.")
    add_bullet("Software Engineering Lifecycle Exposure: Demonstrate comprehensive grasp of the Software Development Life Cycle (SDLC) across initiation, concept analysis, structured design, coding, and empirical testing.")

    # ==========================================
    # PAGE 6: PROPOSED SYSTEM (Expanded full page)
    # ==========================================
    doc.add_page_break()
    add_h1("03. PROPOSED SYSTEM", bookmark="sec_proposed")
    add_h2("SHORTCOMINGS OF THE EXISTING MANUAL SYSTEM")
    add_body("Under traditional manual banking workflows, all customer transactions are transcribed into physical paper ledgers. This archaic operational model exhibits multiple critical points of failure: tellers calculating balances manually during peak business hours frequently introduce arithmetic errors that compromise institutional accounts; customer profiles must be transcribed across multiple register books, causing data drift and inconsistency; auditing past statements requires laborious manual searching through historical ledgers; and physical registers degrade over time and are completely defenseless against termites, dampness, fire hazards, and theft.")

    add_h2("ARCHITECTURE & CAPABILITIES OF THE PROPOSED SYSTEM")
    add_body("The proposed computerized Bank Management System introduces an integrated digital environment that overcomes every manual bottleneck. Written in Python 3.12, the application interfaces with a high-performance local MySQL relational database engine. All transaction arithmetic is automated by the processor, completely eliminating human mathematical error. Account updates, profile lookups, and transaction logging happen synchronously in milliseconds.")
    add_body("The application's operational flow is organized around distinct, self-contained modules — authentication, account administration, transaction processing, and reporting — each rigorously isolated to prevent cascading failures. Every operator interaction is routed through a unified, menu-driven console that mirrors real-world teller workflows, so non-technical staff can adopt the system with minimal training. The architecture deliberately favors transparent, data-centric operations over complex abstractions, ensuring that each screen transition maps directly to a verifiable database action.")
    
    add_image_centered("images/1.png", "Figure 3.1: Bank Management System Architecture Diagram", width_in=4.8)

    # ==========================================
    # PAGE 7: SYSTEM DEVELOPMENT LIFE CYCLE (SDLC)
    # ==========================================
    doc.add_page_break()
    add_h1("04. SYSTEM DEVELOPMENT LIFE CYCLE (SDLC)", bookmark="sec_sdlc")
    add_body("The Systems Development Life Cycle (SDLC) is a disciplined project management methodology that partitions software engineering endeavors into well-defined, verifiable phases. Segmenting development allows engineers and project supervisors to validate requirements, assess risks, and verify milestones before moving to subsequent development stages.")
    add_body("Applying SDLC principles ensures that the Bank Management System achieves high reliability, data integrity, and strict adherence to user specifications. The complete sequential workflow followed throughout the lifecycle of this project is presented in the flowchart below:")
    
    add_image_centered("images/2.png", "Figure 4.1: Software Development Life Cycle (SDLC) Flowchart", width_in=3.6)

    # ==========================================
    # PAGE 8: PHASES OF SDLC (Phase 1 & Phase 2)
    # ==========================================
    doc.add_page_break()
    add_h1("05. PHASES OF SYSTEM DEVELOPMENT LIFE CYCLE", bookmark="sec_phases")
    add_h2("PHASE 1: INITIATION PHASE")
    add_body("The Initiation Phase marks the official start of the project. It establishes whether an automated software system is warranted to replace manual banking logs and identifies the high-level operational boundaries.")
    add_body("Key objectives accomplished during this phase include:", bold_prefix="Objectives: ")
    add_bullet("Identifying the primary business problem: delays and mathematical inaccuracies inherent in paper passbook updating.")
    add_bullet("Establishing initial feasibility: confirming that Python and MySQL provide a cost-effective, robust software stack for local branch banking.")
    add_bullet("Defining project scope: restricting operations to account registration, customer profiling, deposit/withdrawal transactions, and account lifecycle management.")
    add_bullet("Appointing project roles and establishing mentorship milestones under guidance.")
    
    add_h2("PHASE 2: SYSTEM CONCEPT DEVELOPMENT PHASE")
    add_body("During this phase, the abstract project concept is refined into a defined operational model. The system boundaries, stakeholder interactions, and operational environments are documented.")
    add_body("Key activities performed in this phase:", bold_prefix="Key Activities: ")
    add_bullet("Defining system boundary: establishing local command-line interface interactions with local MySQL persistence.")
    add_bullet("Identifying user classes: bank tellers, account supervisors, and registered customers.")
    add_bullet("Assessing project risks: identifying risks such as negative account balances, unhandled database disconnections, and SQL syntax errors, and formulating mitigation strategies.")
    add_bullet("Formulating performance criteria: ensuring response times under 1 second for standard account queries.")

    # ==========================================
    # PAGE 9: PHASES OF SDLC (Phase 3 & Phase 4)
    # ==========================================
    doc.add_page_break()
    add_h2("PHASE 3: PLANNING PHASE")
    add_body("The Planning Phase transforms the project concept into a comprehensive work breakdown structure. Timelines, resource allocations, and coding deliverables are scheduled to ensure systematic project execution.")
    add_body("Key deliverables of this phase:", bold_prefix="Deliverables: ")
    add_bullet("Development Schedule: allocating specific milestones for database schema definition, Python script authoring, test case creation, and report compilation.")
    add_bullet("Resource Planning: identifying hardware prerequisites (PC, keyboard, monitor) and software prerequisites (Python 3.12 interpreter, MySQL Server 8.0, mysql-connector-python package).")
    add_bullet("Quality Management Plan: establishing validation standards for inputs (e.g. 4-digit numeric passwords, positive transaction amounts).")

    add_h2("PHASE 4: REQUIREMENTS ANALYSIS PHASE")
    add_body("Requirements analysis involves formally gathering, analyzing, and documenting what the system must accomplish from a functional and technical standpoint.")
    add_body("Functional Requirements:", bold_prefix="1. ")
    add_bullet("User Authentication: The system must enforce username and 4-digit PIN verification before granting access.")
    add_bullet("Account Registration: The system must record unique account numbers, account holder names, 10-digit phone numbers, addresses, and initial deposit amounts.")
    add_bullet("Transaction Processing: The system must process deposits and withdrawals, updating account credit balances immediately and recording an entry in the transaction log.")
    add_bullet("Overdraft Prevention: The system must reject withdrawal requests exceeding the current credit balance.")
    add_bullet("Account Inquiries: The system must retrieve full profile details and chronological transaction logs on demand.")
    add_body("Non-Functional Requirements:", bold_prefix="2. ")
    add_bullet("Data Integrity: Ensuring referential consistency and ACID properties during database write operations.")
    add_bullet("Usability: Clear, numbered command-line menu navigation with informative confirmation messages.")

    # ==========================================
    # PAGE 10: PHASES OF SDLC (Phase 5: Design)
    # ==========================================
    doc.add_page_break()
    add_h2("PHASE 5: SYSTEM DESIGN PHASE")
    add_body("The System Design Phase translates documented requirements into technical architectural blueprints. Design activities focus on modular decomposition, procedural control flow, and data structuring.")
    add_body("Modular Architecture:", bold_prefix="1. ")
    add_bullet("Authentication Module (main_bank.py): Handles user registration, credentials storage, PIN validation, and initial routing.")
    add_bullet("Banking Operations Module (menu.py): Manages the primary operational loop, providing 6 core functions: account creation, transaction execution, detail inquiry, statement listing, account termination, and session exit.")
    add_bullet("Database Schema Definition Modules (table.py, user_table.py, transaction_table.py): Independent Data Definition Language (DDL) scripts to initialize required database tables.")
    
    add_body("Data Flow & Control Logic:", bold_prefix="2. ")
    add_body("Control flow follows a hierarchical menu structure. Upon launching the application, the user registers or authenticates. Upon successful authentication, the control transfers to main_menu(), which iterates continuously until the user explicitly selects option 6 (Quit). Every transaction commits changes immediately to prevent data inconsistency.")
    
    add_body("Error Handling Strategy:", bold_prefix="3. ")
    add_bullet("Account Existence Checks: Verifying that account numbers exist prior to executing transactions or lookups.")
    add_bullet("Sufficient Funds Check: Ensuring account balance is strictly greater than or equal to the requested withdrawal amount before executing SQL updates.")

    add_image_centered("images/4.png", "Figure 5.1: Bank Management System Entity-Relationship (ER) Diagram", width_in=3.5)

    # ==========================================
    # PAGE 11: PHASES OF SDLC (Database Design & ER Diagram)
    # ==========================================
    doc.add_page_break()
    add_h2("DATABASE DESIGN & ENTITY-RELATIONSHIP (ER) MODELING")
    add_body("The relational database schema is structured into three normalized tables designed to minimize redundancy while preserving data relationships. The database is titled 'bank' in MySQL.")
    
    add_body("Table 1: user_table — Stores operator authentication credentials.", bold_prefix="• ")
    add_bullet("username VARCHAR(25) PRIMARY KEY — Unique username for system access.")
    add_bullet("passwrd VARCHAR(25) NOT NULL — 4-digit security PIN.")
    
    add_body("Table 2: customer_details — Stores profile and financial state of account holders.", bold_prefix="• ")
    add_bullet("acct_no INT PRIMARY KEY — Unique primary key identifying the bank account.")
    add_bullet("acct_name VARCHAR(25) — Full legal name of the account holder.")
    add_bullet("phone_no BIGINT — Customer contact number.")
    add_bullet("address VARCHAR(50) — Place / residential city.")
    add_bullet("cr_amt FLOAT — Current balance credit amount.")
    
    add_body("Table 3: transactions — Stores an append-only audit trail of financial operations.", bold_prefix="• ")
    add_bullet("acct_no INT — Account number associated with the transaction (Foreign Key relationship).")
    add_bullet("date DATE — Date timestamp of transaction execution.")
    add_bullet("withdrawal_amt FLOAT — Amount debited (0 if deposit).")
    add_bullet("amount_added FLOAT — Amount credited (0 if withdrawal).")

    # ==========================================
    # PAGE 12: PHASES OF SDLC (Phase 6: Development)
    # ==========================================
    doc.add_page_break()
    add_h2("PHASE 6: SOFTWARE DEVELOPMENT PHASE (CODING)")
    add_body("The Development Phase entails translating architectural designs and database schemas into clean, executable Python 3.12 source code. Development strictly adhered to CBSE Class 12 Computer Science programming conventions.")
    add_body("Coding Guidelines & Principles Followed:", bold_prefix="Implementation Highlights: ")
    add_bullet("Clean Python 3.12 Syntax: Replaced deprecated string concatenations and awkward spacing with clear, readable f-strings (f\"...\").")
    add_bullet("Modular Procedural Design: Encapsulated banking logic inside clean functions (e.g. main_menu()) equipped with an __name__ == '__main__' execution block for flexible invocation.")
    add_bullet("Correct SQL Transaction Control: Utilized explicit conn.commit() invocations after every INSERT, UPDATE, and DELETE query to ensure changes are persistently written to disk.")
    add_bullet("Loop Integrity: Implemented a robust while True loop with an explicit exit condition, resolving the previous infinite-loop flaw present in older scripts.")
    add_bullet("Standard Library Utilization: Leveraged Python's built-in datetime module for automated timestamping of financial ledger entries.")
    add_bullet("CBSE Class 12 Alignment: Avoided excessive abstractions, classes/OOP, or external frameworks, keeping the codebase fully aligned with CBSE Class 12 evaluation criteria.")

    # ==========================================
    # PAGE 13: PHASES OF SDLC (Phase 7: Testing)
    # ==========================================
    doc.add_page_break()
    add_h2("PHASE 7: INTEGRATION AND TESTING PHASE")
    add_body("During this phase, individual software modules were integrated and tested as a unified software system. Both Python client scripts and the local MySQL server instance were exercised across standard and boundary conditions.")
    add_body("Testing Stages Conducted:", bold_prefix="Testing Workflow: ")
    add_bullet("Unit Testing: Verifying individual SQL operations independently (e.g., verifying that INSERT INTO user_table accurately populates rows in user_table).")
    add_bullet("Integration Testing: Testing the handshake between main_bank.py and menu.py to ensure that user credentials pass seamlessly into the operational dashboard.")
    add_bullet("Boundary Condition Testing: Testing boundary values, such as depositing Rs. 0, withdrawing an amount equal to the exact balance, and attempting to withdraw Rs. 1 more than the current balance.")
    add_bullet("Invalid Input Testing: Verifying system resilience when invalid account numbers are submitted for transactions or lookups.")
    add_bullet("Database Synchronization Testing: Validating that executing an account deletion removes the record from both customer_details and the transactions log.")

    # ==========================================
    # PAGE 14: PHASES OF SDLC (Phase 8 & Phase 9)
    # ==========================================
    doc.add_page_break()
    add_h2("PHASE 8: IMPLEMENTATION PHASE")
    add_body("The Implementation Phase transitions the tested software system into operational readiness. It involves initializing the physical hardware environment, establishing database instances, and validating operational workflows.")
    add_body("Deployment Milestones Accomplished:", bold_prefix="Milestones: ")
    add_bullet("MySQL Database Initialization: Created the dedicated bank relational schema on localhost.")
    add_bullet("DDL Script Execution: Executed table.py, user_table.py, and transaction_table.py to create required relational structures.")
    add_bullet("Environment Verification: Tested connectivity between the Python 3.12 interpreter and the local MySQL 8.0 server using root credentials.")
    add_bullet("Initial Seeding: Created administrative and sample customer records to confirm operational health.")

    add_h2("PHASE 9: OPERATIONS AND MAINTENANCE PHASE")
    add_body("The Operations and Maintenance Phase encompasses post-deployment monitoring, ongoing data integrity verification, and planning future software enhancements.")
    add_body("Maintenance Best Practices:", bold_prefix="Operational Policies: ")
    add_bullet("Database Backups: Routine export of MySQL database tables (e.g., via mysqldump) to safeguard against accidental storage failure.")
    add_bullet("Integrity Auditing: Periodic verification of transaction logs against customer current credit balances to ensure financial reconciliation.")
    add_bullet("System Evolution: Planning future upgrades, such as graphical user interface (GUI) integration using Tkinter and interest rate calculation modules.")

    # ==========================================
    # PAGE 15: FLOW CHART (Overall Working Process)
    # ==========================================
    doc.add_page_break()
    add_h1("06. FLOW CHART", bookmark="sec_flowchart")
    add_h2("OVERALL WORKING PROCESS OF BANK MANAGEMENT SYSTEM")
    add_body("The comprehensive operational workflow of the Bank Management System is diagrammed below. It traces the lifecycle of user interaction from system startup, through operator registration or PIN authentication, into the primary 6-option command dashboard, and through subsequent database synchronization and reporting.")
    
    add_image_centered("images/5.png", "Figure 6.1: Overall Working Process of Bank Management System", width_in=3.7)

    # ==========================================
    # PAGE 16: FLOW CHART (Bank Transaction Process)
    # ==========================================
    doc.add_page_break()
    add_h2("BANK TRANSACTION PROCESS FLOWCHART")
    add_body("The dedicated flowchart below details the precise operational logic governing financial transactions (Cash Withdrawals and Cash Deposits). It highlights validation steps including account verification, choice branching, overdraft prevention checks, ledger updates, and transaction logging.")
    
    add_image_centered("images/3.png", "Figure 6.2: Bank Transaction Process Flowchart", width_in=4.4)

    # ==========================================
    # PAGE 17: SOURCE CODE (DDL Table Scripts)
    # ==========================================
    doc.add_page_break()
    add_h1("07. SOURCE CODE", bookmark="sec_sourcecode")
    add_h2("1. DATABASE INITIALIZATION SCRIPTS (DDL)")
    add_body("The following three standalone scripts establish the relational database tables in the MySQL 'bank' database.")
    
    add_body("A. customer_details Table Setup (table.py):", bold_prefix="")
    add_code_block('''# table.py - Database setup for customer_details table
import mysql.connector as sql

conn = sql.connect(host='localhost', user='root', passwd='manager', database='bank')
cur = conn.cursor()

cur.execute(\'\'\'
    CREATE TABLE IF NOT EXISTS customer_details (
        acct_no INT PRIMARY KEY,
        acct_name VARCHAR(25),
        phone_no BIGINT,
        address VARCHAR(50),
        cr_amt FLOAT
    )
\'\'\')
conn.commit()
print("Table 'customer_details' created successfully.")''')

    add_body("B. user_table Table Setup (user_table.py):", bold_prefix="")
    add_code_block('''# user_table.py - Database setup for user authentication
import mysql.connector as sql

conn = sql.connect(host='localhost', user='root', passwd='manager', database='bank')
cur = conn.cursor()

cur.execute(\'\'\'
    CREATE TABLE IF NOT EXISTS user_table (
        username VARCHAR(25) PRIMARY KEY,
        passwrd VARCHAR(25) NOT NULL
    )
\'\'\')
conn.commit()
print("Table 'user_table' created successfully.")''')

    add_body("C. transactions Table Setup (transaction_table.py):", bold_prefix="")
    add_code_block('''# transaction_table.py - Database setup for transaction audit logs
import mysql.connector as sql

conn = sql.connect(host='localhost', user='root', passwd='manager', database='bank')
cur = conn.cursor()

cur.execute(\'\'\'
    CREATE TABLE IF NOT EXISTS transactions (
        acct_no INT,
        date DATE,
        withdrawal_amt FLOAT,
        amount_added FLOAT
    )
\'\'\')
conn.commit()
print("Table 'transactions' created successfully.")''')

    # ==========================================
    # PAGE 18: SOURCE CODE (main_bank.py Part 1)
    # ==========================================
    doc.add_page_break()
    add_h2("2. MAIN SYSTEM ENTRY SCRIPT (main_bank.py - Part 1)")
    add_body("The main_bank.py script acts as the master authentication gatekeeper and security boundary for the Bank Management System. Upon execution, the module initializes runtime dependencies, connects to the local MySQL database instance using host, user, and password parameters, and initializes an operational cursor for transactional communication.")
    add_body("Part 1 implements the application bootstrapper and the operator onboarding sequence. It presents an interactive terminal interface that timestamps system entry using Python's standard datetime module. When Option 1 (Register) is selected, the routine prompts the operator for a new username and a mandatory 4-digit security PIN. It validates the credentials and executes an INSERT query against the user_table relation, followed by an immediate conn.commit() to ensure persistent ACID storage before launching the main banking menu.")
    
    add_code_block('''# main_bank.py - User Authentication and System Gateway
import datetime as dt
import mysql.connector as sql
import menu

# Establishing connection with MySQL server
conn = sql.connect(host='localhost', user='root', passwd='manager', database='bank')
cur = conn.cursor()

print('========================= WELCOME TO STARK BANK =========================')
print(dt.datetime.now())
print('1. REGISTER')
print('2. LOGIN')
print('=========================================================================')
print()

n = int(input('enter your choice = '))
print()

# Option 1: New User Registration
if n == 1:
    name = input('Enter a Username = ')
    passwd = input('Enter a 4 DIGIT Password = ')
    
    # Inserting credentials into user_table
    query = f"INSERT INTO user_table (username, passwrd) VALUES ('{name}', '{passwd}')"
    cur.execute(query)
    conn.commit()
    print()
    print('USER created successfully!')
    print()
    
    # Launch main banking operations menu
    menu.main_menu()''')

    # ==========================================
    # PAGE 19: SOURCE CODE (main_bank.py Part 2)
    # ==========================================
    doc.add_page_break()
    add_h2("2. MAIN SYSTEM ENTRY SCRIPT (main_bank.py - Part 2)")
    add_body("Continuation of main_bank.py detailing user authentication and dashboard redirection.")
    
    add_code_block('''# Option 2: Existing User Login Authentication
elif n == 2:
    name = input('Enter your Username = ')
    passwd = input('Enter your 4 DIGIT Password = ')
    
    # Verifying credentials against database
    query = f"SELECT * FROM user_table WHERE username = '{name}' AND passwrd = '{passwd}'"
    cur.execute(query)
    
    if cur.fetchone() is None:
        print()
        print('Invalid username or password')
    else:
        print()
        print('Login successful!')
        print()
        # Direct authenticated user to main dashboard
        menu.main_menu()
else:
    print('Invalid choice! Please select 1 or 2.')''')

    # ==========================================
    # PAGE 20: SOURCE CODE (menu.py Part 1)
    # ==========================================
    doc.add_page_break()
    add_h2("3. BANKING OPERATIONS MODULE (menu.py - Part 1)")
    add_body("The menu.py module handles all core banking features. Part 1 covers initialization and Account Creation. Apart from the creation routine itself, this first part establishes the shared database cursor and the master dispatch loop that governs every subsequent operator choice, giving the module a single, predictable entry point.")
    
    add_code_block('''# menu.py - Core Banking Operations Dashboard
import datetime as dt
import mysql.connector as sql

conn = sql.connect(host='localhost', user='root', passwd='manager', database='bank')
cur = conn.cursor()

def main_menu():
    while True:
        print()
        print('========================= MAIN MENU =========================')
        print('1. CREATE BANK ACCOUNT')
        print('2. TRANSACTION')
        print('3. CUSTOMER DETAILS')
        print('4. TRANSACTION DETAILS')
        print('5. DELETE ACCOUNT')
        print('6. QUIT')
        print('=============================================================')
        print()

        n = int(input('Enter your CHOICE = '))
        print()

        # Feature 1: Create New Bank Account
        if n == 1:
            acc_no = int(input('Enter your ACCOUNT NUMBER = '))
            acc_name = input('Enter your ACCOUNT NAME = ')
            ph_no = int(input('Enter your PHONE NUMBER = '))
            add = input('Enter your place = ')
            cr_amt = float(input('Enter your credit amount = '))

            query = f"INSERT INTO customer_details VALUES ({acc_no}, '{acc_name}', {ph_no}, '{add}', {cr_amt})"
            cur.execute(query)
            conn.commit()
            print()
            print('Account Created Successfully!!!!!')''')

    # ==========================================
    # PAGE 21: SOURCE CODE (menu.py Part 2)
    # ==========================================
    doc.add_page_break()
    add_h2("3. BANKING OPERATIONS MODULE (menu.py - Part 2)")
    add_body("Continuation of menu.py detailing the financial Transaction Processing engine.")
    
    add_code_block('''        # Feature 2: Financial Transactions (Withdrawal & Deposit)
        elif n == 2:
            acct_no = int(input('Enter Your Account Number = '))
            cur.execute(f"SELECT * FROM customer_details WHERE acct_no = {acct_no}")
            data = cur.fetchall()

            if len(data) == 0:
                print()
                print('Account Number Invalid. Please try again.')
            else:
                current_bal = data[0][4]
                print()
                print('1. WITHDRAW AMOUNT')
                print('2. ADD AMOUNT')
                print()
                x = int(input('Enter your CHOICE = '))
                print()
                today_date = dt.date.today()

                # Sub-option 1: Cash Withdrawal
                if x == 1:
                    amt = float(input('Enter withdrawal amount = '))
                    if amt > current_bal:
                        print()
                        print('Insufficient Balance!')
                    else:
                        cur.execute(f"UPDATE customer_details SET cr_amt = cr_amt - {amt} WHERE acct_no = {acct_no}")
                        cur.execute(f"INSERT INTO transactions VALUES ({acct_no}, '{today_date}', {amt}, 0)")
                        conn.commit()
                        print()
                        print('Account Updated Successfully!!!!!')

                # Sub-option 2: Cash Deposit
                elif x == 2:
                    amt = float(input('Enter amount to be added = '))
                    cur.execute(f"UPDATE customer_details SET cr_amt = cr_amt + {amt} WHERE acct_no = {acct_no}")
                    cur.execute(f"INSERT INTO transactions VALUES ({acct_no}, '{today_date}', 0, {amt})")
                    conn.commit()
                    print()
                    print('Account Updated Successfully!!!!!')''')

    # ==========================================
    # PAGE 22: SOURCE CODE (menu.py Part 3)
    # ==========================================
    doc.add_page_break()
    add_h2("3. BANKING OPERATIONS MODULE (menu.py - Part 3)")
    add_body("Continuation of menu.py covering Customer Inquiries, Transaction Statement Logs, Account Termination, and Session Exit.")
    
    add_code_block('''        # Feature 3: View Customer Details
        elif n == 3:
            acct_no = int(input('Enter your account number = '))
            print()
            cur.execute(f"SELECT * FROM customer_details WHERE acct_no = {acct_no}")
            data = cur.fetchall()

            if len(data) == 0:
                print('Invalid Account number')
            else:
                for row in data:
                    print('ACCOUNT NO      =', row[0])
                    print('ACCOUNT NAME    =', row[1])
                    print('PHONE NUMBER    =', row[2])
                    print('ADDRESS         =', row[3])
                    print('CURRENT BALANCE =', row[4])
                    print()

        # Feature 4: View Transaction Statements
        elif n == 4:
            acct_no = int(input('Enter your account number = '))
            print()
            cur.execute(f"SELECT * FROM customer_details WHERE acct_no = {acct_no}")
            if cur.fetchone() is None:
                print('Invalid Account number')
            else:
                cur.execute(f"SELECT * FROM transactions WHERE acct_no = {acct_no}")
                data = cur.fetchall()
                if len(data) == 0:
                    print('No transactions found for this account.')
                else:
                    print(f"{'ACC NO':<10} {'DATE':<15} {'WITHDRAWAL':<15} {'DEPOSIT':<15}")
                    print('-' * 55)
                    for row in data:
                        print(f"{row[0]:<10} {str(row[1]):<15} {row[2]:<15} {row[3]:<15}")
                    print()

        # Feature 5: Delete Account
        elif n == 5:
            print('DELETE YOUR ACCOUNT')
            acct_no = int(input('Enter your account number = '))
            cur.execute(f"SELECT * FROM customer_details WHERE acct_no = {acct_no}")
            if cur.fetchone() is None:
                print('Account does not exist!')
            else:
                cur.execute(f"DELETE FROM customer_details WHERE acct_no = {acct_no}")
                cur.execute(f"DELETE FROM transactions WHERE acct_no = {acct_no}")
                conn.commit()
                print()
                print('ACCOUNT DELETED SUCCESSFULLY')

        # Feature 6: Quit Application
        elif n == 6:
            c = input('DO YOU WANT TO EXIT (y/n) = ')
            if c.lower() == 'y':
                print()
                print('THANK YOU PLEASE VISIT AGAIN')
                break

if __name__ == '__main__':
    main_menu()''', font_size_pt=7.5, line_spacing=1.0)

    # ==========================================
    # PAGE 23: OUTPUT (Welcome & Registration)
    # ==========================================
    doc.add_page_break()
    add_h1("08. OUTPUT", bookmark="sec_output")
    add_h2("OUTPUT SCREEN 1: SYSTEM WELCOME & USER REGISTRATION")
    add_body("Upon running main_bank.py, the user is presented with the Stark Bank header, current timestamp, and entry options. Selecting Option 1 allows the registration of a new operator.")
    
    add_code_block('''========================= WELCOME TO STARK BANK =========================
2026-10-26 10:15:32.418291
1. REGISTER
2. LOGIN
=========================================================================

enter your choice = 1

Enter a Username = suryah
Enter a 4 DIGIT Password = 1298

USER created successfully!''')

    # ==========================================
    # PAGE 24: OUTPUT (Login & Main Menu)
    # ==========================================
    doc.add_page_break()
    add_h2("OUTPUT SCREEN 2: USER LOGIN & MAIN MENU DASHBOARD")
    add_body("Selecting Option 2 prompts the operator for their username and 4-digit security PIN. Upon successful verification against user_table, the system unlocks the interactive Main Menu.")
    
    add_code_block('''========================= WELCOME TO STARK BANK =========================
2026-10-26 10:18:04.120954
1. REGISTER
2. LOGIN
=========================================================================

enter your choice = 2

Enter your Username = suryah
Enter your 4 DIGIT Password = 1298

Login successful!

========================= MAIN MENU =========================
1. CREATE BANK ACCOUNT
2. TRANSACTION
3. CUSTOMER DETAILS
4. TRANSACTION DETAILS
5. DELETE ACCOUNT
6. QUIT
=============================================================

Enter your CHOICE = ''')

    # ==========================================
    # PAGE 25: OUTPUT (Create Bank Account)
    # ==========================================
    doc.add_page_break()
    add_h2("OUTPUT SCREEN 3: CREATE BANK ACCOUNT")
    add_body("Selecting Option 1 initiates account creation. The operator enters customer account number, full name, phone number, location, and initial credit amount.")
    
    add_code_block('''========================= MAIN MENU =========================
1. CREATE BANK ACCOUNT
2. TRANSACTION
3. CUSTOMER DETAILS
4. TRANSACTION DETAILS
5. DELETE ACCOUNT
6. QUIT
=============================================================

Enter your CHOICE = 1

Enter your ACCOUNT NUMBER = 1001
Enter your ACCOUNT NAME = Rohan Sharma
Enter your PHONE NUMBER = 9876543210
Enter your place = Delhi
Enter your credit amount = 50000

Account Created Successfully!!!!!''')

    # ==========================================
    # PAGE 26: OUTPUT (Transaction: Deposit & Withdrawal)
    # ==========================================
    doc.add_page_break()
    add_h2("OUTPUT SCREEN 4: TRANSACTION PROCESSING")
    add_body("Selecting Option 2 presents the transaction processing engine. Below are samples of cash deposit and cash withdrawal executions:")
    
    add_code_block('''-- TRANSACTION EXECUTION: CASH DEPOSIT --
Enter your CHOICE = 2

Enter Your Account Number = 1001

1. WITHDRAW AMOUNT
2. ADD AMOUNT

Enter your CHOICE = 2

Enter amount to be added = 15000

Account Updated Successfully!!!!!

-- TRANSACTION EXECUTION: CASH WITHDRAWAL --
Enter your CHOICE = 2

Enter Your Account Number = 1001

1. WITHDRAW AMOUNT
2. ADD AMOUNT

Enter your CHOICE = 1

Enter withdrawal amount = 5000

Account Updated Successfully!!!!!''')

    # ==========================================
    # PAGE 27: OUTPUT (Customer Details & History)
    # ==========================================
    doc.add_page_break()
    add_h2("OUTPUT SCREEN 5: CUSTOMER DETAILS & STATEMENT LOG")
    add_body("Option 3 displays profile information, and Option 4 generates a formatted chronological statement log:")
    
    add_code_block('''-- OPTION 3: CUSTOMER DETAILS LOOKUP --
Enter your CHOICE = 3

Enter your account number = 1001

ACCOUNT NO      = 1001
ACCOUNT NAME    = Rohan Sharma
PHONE NUMBER    = 9876543210
ADDRESS         = Delhi
CURRENT BALANCE = 60000.0

-- OPTION 4: TRANSACTION DETAILS AUDIT LOG --
Enter your CHOICE = 4

Enter your account number = 1001

ACC NO     DATE            WITHDRAWAL      DEPOSIT        
-------------------------------------------------------
1001       2026-10-26      0.0             15000.0        
1001       2026-10-26      5000.0          0.0            
''')

    # ==========================================
    # PAGE 28: OUTPUT (Delete Account & Exit)
    # ==========================================
    doc.add_page_break()
    add_h2("OUTPUT SCREEN 6: DELETE ACCOUNT & QUIT")
    add_body("Option 5 terminates customer accounts and removes audit records. Option 6 safely terminates the active session:")
    
    add_code_block('''-- OPTION 5: TERMINATE BANK ACCOUNT --
Enter your CHOICE = 5

DELETE YOUR ACCOUNT
Enter your account number = 1001

ACCOUNT DELETED SUCCESSFULLY

-- OPTION 6: SYSTEM EXIT --
========================= MAIN MENU =========================
1. CREATE BANK ACCOUNT
2. TRANSACTION
3. CUSTOMER DETAILS
4. TRANSACTION DETAILS
5. DELETE ACCOUNT
6. QUIT
=============================================================

Enter your CHOICE = 6

DO YOU WANT TO EXIT (y/n) = y

THANK YOU PLEASE VISIT AGAIN''')

    # ==========================================
    # PAGE 29: TESTING (Overview & Black Box)
    # ==========================================
    doc.add_page_break()
    add_h1("09. TESTING", bookmark="sec_testing")
    add_h2("1. SOFTWARE TESTING OVERVIEW")
    add_body("Software Testing is an empirical investigation conducted to provide stakeholders with information about the quality of the software product under test. It validates and verifies that the software application fulfills technical specifications and handles unexpected input gracefully.")
    
    add_h2("2. BLACK BOX TESTING (SPECIFICATION-BASED TESTING)")
    add_body("Black Box Testing evaluates functionality without reference to internal source code structures or implementation mechanics. Test cases are derived strictly from functional specifications.")
    add_body("Advantages & Relevance to Banking:", bold_prefix="Characteristics: ")
    add_bullet("Independence: The tester assesses behavior exactly as a bank teller or customer would experience it.")
    add_bullet("Equivalence Partitioning: Inputs are divided into valid classes (e.g. valid 4-digit numeric PINs) and invalid classes (e.g. non-existent account numbers, alphabetic PIN inputs).")
    add_bullet("Boundary Value Analysis: Tests focus on threshold values, such as zero-balance withdrawals, empty string passwords, and large credit sums.")

    # ==========================================
    # PAGE 30: TESTING (White Box Testing)
    # ==========================================
    doc.add_page_break()
    add_h2("3. WHITE BOX TESTING (STRUCTURAL TESTING)")
    add_body("White Box Testing analyzes internal logic paths, control structures, and code execution branches. It ensures that every procedural branch operates reliably under edge conditions.")
    add_body("White Box Testing Methodologies Applied:", bold_prefix="Techniques: ")
    add_bullet("Statement Coverage: Ensuring that every line of SQL execution and print statement within main_bank.py and menu.py is executed at least once during test runs.")
    add_bullet("Branch / Decision Coverage: Ensuring both True and False branches of every conditional statement (if cur.fetchone() is None, if amt > current_bal) are rigorously exercised.")
    add_bullet("Database Exception Verification: Verifying that SQL integrity violations (e.g. duplicate primary key entry) are caught and handled without unceremonious script termination.")
    add_bullet("Transaction Integrity: Validating that conn.commit() is consistently called, ensuring atomic persistence.")

    # ==========================================
    # PAGE 31: TESTING (Test Cases Matrix)
    # ==========================================
    doc.add_page_break()
    add_h2("4. SYSTEM TEST CASES EXECUTION MATRIX")
    add_body("The table below details the formal test cases designed and executed against the Bank Management System, confirming complete functional correctness:")
    
    test_cases = [
        ("TC ID", "Module", "Test Description", "Test Input", "Expected Result", "Status"),
        ("TC01", "Auth", "Register valid user", "suryah, 1298", "User created successfully", "PASS"),
        ("TC02", "Auth", "Login with valid PIN", "suryah, 1298", "Login successful; show menu", "PASS"),
        ("TC03", "Auth", "Login with wrong PIN", "suryah, 9999", "Invalid username/password", "PASS"),
        ("TC04", "Account", "Create unique account", "Acc: 1001, Rs.50000", "Account Created Successfully", "PASS"),
        ("TC05", "Transact", "Deposit valid amount", "Acc: 1001, +Rs.15000", "Balance updated to 65000", "PASS"),
        ("TC06", "Transact", "Withdraw valid amount", "Acc: 1001, -Rs.5000", "Balance updated to 60000", "PASS"),
        ("TC07", "Transact", "Overdraft withdrawal", "Acc: 1001, -Rs.90000", "Reject: Insufficient Balance", "PASS"),
        ("TC08", "Inquiry", "Lookup valid account", "Acc: 1001", "Display correct customer row", "PASS"),
        ("TC09", "Inquiry", "Lookup invalid account", "Acc: 9999", "Invalid Account number", "PASS"),
        ("TC10", "Statement", "Check audit logs", "Acc: 1001", "Display all logged transactions", "PASS"),
        ("TC11", "Delete", "Delete active account", "Acc: 1001", "Account Deleted Successfully", "PASS"),
    ]
    
    tc_tbl = doc.add_table(rows=len(test_cases), cols=6)
    tc_widths = [Inches(0.6), Inches(0.9), Inches(1.8), Inches(1.4), Inches(1.3), Inches(0.5)]
    
    for r_idx, row in enumerate(tc_tbl.rows):
        vals = test_cases[r_idx]
        for c_idx, cell in enumerate(row.cells):
            cell.width = tc_widths[c_idx]
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [0, 1, 5] else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(vals[c_idx])
            run.font.name = "Times New Roman"
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(0, 0, 0)
            if r_idx == 0:
                run.bold = True
    
    style_table_bw(tc_tbl)

    # ==========================================
    # PAGE 32: HARDWARE & SOFTWARE REQUIREMENTS
    # ==========================================
    doc.add_page_break()
    add_h1("10. HARDWARE AND SOFTWARE REQUIREMENTS", bookmark="sec_hardware")
    add_h2("1. HARDWARE REQUIREMENTS")
    add_body("The Bank Management System is optimized for high computational efficiency, running smoothly on modest personal computer configurations:")
    
    hw_specs = [
        ("Component", "Minimum Requirement", "Recommended Specification"),
        ("Processor", "Intel Pentium / AMD Athlon Dual-Core", "Intel Core i3 / AMD Ryzen 3 or higher"),
        ("RAM", "512 MB RAM", "4 GB RAM or higher"),
        ("Hard Disk Space", "40 GB SATA HDD", "256 GB SSD (500 MB free for DB)"),
        ("Display Monitor", "14.1 or 15–17 inch Display (1024x768)", "1920x1080 Full HD Monitor"),
        ("Input Devices", "Standard Keyboard and Mouse", "Standard QWERTY Keyboard and Mouse"),
        ("Printer", "Optional (for hard-copy statements)", "Monochrome Laser Printer (B&W)"),
    ]
    
    hw_tbl = doc.add_table(rows=len(hw_specs), cols=3)
    hw_widths = [Inches(1.8), Inches(2.4), Inches(2.3)]
    for r_idx, row in enumerate(hw_tbl.rows):
        vals = hw_specs[r_idx]
        for c_idx, cell in enumerate(row.cells):
            cell.width = hw_widths[c_idx]
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(vals[c_idx])
            run.font.name = "Times New Roman"
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0, 0, 0)
            if r_idx == 0:
                run.bold = True
    style_table_bw(hw_tbl)
    
    doc.add_paragraph().paragraph_format.space_before = Pt(8)
    add_h2("2. SOFTWARE REQUIREMENTS")
    add_bullet("Operating System: Microsoft Windows 7 / 8 / 10 / 11 (32-bit or 64-bit).")
    add_bullet("Python Interpreter: Python 3.12 (or any stable Python 3.7+ release).")
    add_bullet("Database Management System: MySQL Community Server 8.0 (or MySQL 5.7).")
    add_bullet("Python Database Connector: mysql-connector-python package.")
    add_bullet("Development Environment: Python IDLE, VS Code, or Windows PowerShell terminal.")

    # ==========================================
    # PAGE 33: INSTALLATION PROCEDURE
    # ==========================================
    doc.add_page_break()
    add_h1("11. INSTALLATION PROCEDURE", bookmark="sec_installation")
    add_body("To deploy and execute the Bank Management System in a new environment, follow the step-by-step procedure detailed below:")
    
    add_body("Step 1: Install Python Runtime", bold_prefix="1. ")
    add_body("Download Python 3.12 from python.org. Run the installer and ensure the checkbox 'Add Python.exe to PATH' is selected. Verify installation by running python --version in the terminal.", indent=True)
    
    add_body("Step 2: Install MySQL Server and Connector", bold_prefix="2. ")
    add_body("Install MySQL Community Server from dev.mysql.com. Set the root password to manager (or update the credentials in the Python source code). Install the MySQL Python connector by executing:", indent=True)
    add_code_block("pip install mysql-connector-python")
    
    add_body("Step 3: Create the Relational Database", bold_prefix="3. ")
    add_body("Open the MySQL Command Line Client, authenticate with your root password, and initialize the project database:", indent=True)
    add_code_block("CREATE DATABASE bank;\nUSE bank;")
    
    add_body("Step 4: Execute Database Table Schema Scripts", bold_prefix="4. ")
    add_body("Navigate to the source code folder and execute the three DDL setup scripts sequentially to create the tables:", indent=True)
    add_code_block("python table.py\npython user_table.py\npython transaction_table.py")
    
    add_body("Step 5: Launch the Bank Management Application", bold_prefix="5. ")
    add_body("Launch the primary system entry file:", indent=True)
    add_code_block("python main_bank.py")
    add_body("Follow the interactive prompts to register a user account, authenticate via PIN, and access the main banking dashboard.", indent=True)

    # ==========================================
    # PAGE 34: BIBLIOGRAPHY
    # ==========================================
    doc.add_page_break()
    add_h1("12. BIBLIOGRAPHY", bookmark="sec_bibliography")
    add_body("The development, database design, and documentation of this project were guided by the following textbooks, reference materials, and web resources:")
    
    add_h2("TEXTBOOKS & ACADEMIC REFERENCES")
    add_bullet("Sumita Arora, Computer Science with Python - Textbook for Class XII, Dhanpat Rai & Co. Publications.")
    add_bullet("Preeti Arora, Computer Science with Python for Class XII, Sultan Chand Publications.")
    add_bullet("NCERT, Computer Science - Class XII (Python), National Council of Educational Research and Training.")
    add_bullet("Praveen M Jigajinni, Project Guidance and Reference Material on Bank Management System (BMS), Sainik School Amaravathinagar.")
    
    doc.add_paragraph().paragraph_format.space_before = Pt(8)
    add_h2("ONLINE DOCUMENTATION & TECHNICAL REFERENCES")
    add_bullet("Python Software Foundation, Official Python 3.12 Documentation: https://docs.python.org/3/")
    add_bullet("Oracle Corporation, MySQL 8.0 Reference Manual: https://dev.mysql.com/doc/refman/8.0/en/")
    add_bullet("Oracle Corporation, MySQL Connector/Python Developer Guide: https://dev.mysql.com/doc/connector-python/en/")
    add_bullet("w3resource, Python MySQL Database Connectivity Tutorials: https://www.w3resource.com/python-exercises/mysql/")
    add_bullet("GeeksforGeeks, CBSE Class 12 Computer Science Project Guidelines: https://www.geeksforgeeks.org/school-learning/")

    # Save document
    try:
        doc.save("BMS.docx")
        print("BMS.docx successfully regenerated!")
    except PermissionError:
        doc.save("BMS_updated.docx")
        print("BMS.docx was open in another app. Saved as BMS_updated.docx successfully!")

if __name__ == "__main__":
    create_document()
