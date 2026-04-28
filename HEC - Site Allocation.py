#!/usr/bin/env python
# coding: utf-8

# In[1]:


import sys
from datetime import date, timedelta

from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor, QPixmap, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QSplitter, QTreeWidget, QTreeWidgetItem, QTableWidget,
    QTableWidgetItem, QPushButton, QLabel, QComboBox,
    QLineEdit, QDateEdit, QMessageBox, QTabWidget,
    QListWidget, QDockWidget, QMenu, QInputDialog, QListWidgetItem,
    QGroupBox
)

import requests
import json
import os

class SchedulerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HEC QHSE - Site Operations Scheduler")
        self.resize(1600, 900)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #012f63;
            }
            
            QTabWidget::pane {
                border: none;
                background: white;
            }
            
            QTabBar::tab {
                background: #d9e3f0;
                color: black;
                padding: 10px 20px;
                min-width: 120px;
                min-height: 25px;
                border-radius: 8px;
                margin-right: 5px;
            }
            
            QTabBar::tab:selected {
                background: #012f63;
                color: white;
                font-weight: bold;
            }
            
            QTableWidget {
                alternate-background-color: #f7f9fc;
            }
            
            QPushButton {
                background-color: #012f63;
                color: white;
                padding: 8px;
                border-radius: 6px;
                font-weight: bold;
            }
            
            QPushButton:hover {
                background-color: #014a99;
            }
            
            QTreeWidget, QTableWidget, QListWidget {
                background: white;
                border: 1px solid #dcdcdc;
                gridline-color: #cccccc;
            }
            
            QHeaderView::section {
                background-color: #d9e3f0;
                color: #1f2d3d;
                padding: 8px;
                border: 2px solid white;
                border-radius: 12px;
                font-weight: bold;
            }
            
            QGroupBox {
                background: white;
                border: 1px solid #dcdcdc;
                border-radius: 8px;
                margin-top: 10px;
                font-weight: bold;
                padding: 10px;
            }
            
            QLineEdit, QComboBox, QDateEdit {
                padding: 5px;
                border: 1px solid #ccc;
                border-radius: 5px;
                background: white;
            }
            
            QScrollBar:horizontal {
                background: #e6ebf2;
                height: 14px;
                border-radius: 7px;
                margin: 2px;
            }

            QScrollBar::handle:horizontal {
                background: #7d8ca3;
                border-radius: 7px;
                min-width: 40px;
            }

            QScrollBar::handle:horizontal:hover {
                background: #5f728f;
            }

            QScrollBar::add-line:horizontal,
            QScrollBar::sub-line:horizontal {
                border: none;
                background: none;
            }

            QScrollBar::add-page:horizontal,
            QScrollBar::sub-page:horizontal {
                background: none;
            }
            
        """)
        self.load_window_icon()

        #'self.projects = {
         #   "HOW03": ["CLO"],
         #   "OW3": ["Drums Unloading", "Sheath Test Rostock"],
        #    "Oresund2": ["LJB Jointing"],
        #}

        #self.employees = [
        #    "APB", "THS", "PAP", "LER", "KOK", "GIK", "KOB"
        #]
        
        self.projects = {}
        self.employees = []
        self.site_periods = {}
        self.site_to_row = {}
        self.assignments = {}
        self.employee_leaves = {}

        self.load_data()

        self.init_ui()

    # -----------------------------
    # MAIN UI
    # -----------------------------
    def load_data(self):
        if os.path.exists("scheduler_data.json"):

            with open("scheduler_data.json", "r") as f:
                data = json.load(f)

            self.projects = data.get(
                "projects",
                {}
            )

            self.employees = data.get(
                "employees",
                []
            )

            # ------------------------
            # restore site periods
            # ------------------------
            self.site_periods = {}

            raw_periods = data.get(
                "site_periods",
                {}
            )

            for key, period in raw_periods.items():
                project, site = key.split("|")

                self.site_periods[
                    (project, site)
                ] = (
                    date.fromisoformat(
                        period["start"]
                    ),
                    date.fromisoformat(
                        period["end"]
                    )
                )

            # ------------------------
            # restore assignments
            # ------------------------
            self.assignments = {}

            raw_assignments = data.get(
                "assignments",
                {}
            )

            for key, employee in raw_assignments.items():
                project, site, assigned_date = key.split("|")

                self.assignments[
                    (
                        project,
                        site,
                        date.fromisoformat(
                            assigned_date
                        )
                    )
                ] = employee
                
            self.employee_leaves = {}

            raw_leaves = data.get(
                "employee_leaves",
                {}
            )

            for employee, leaves in raw_leaves.items():

                self.employee_leaves[employee] = []

                for leave in leaves:
                    self.employee_leaves[employee].append(
                        (
                            date.fromisoformat(
                                leave["start"]
                            ),
                            date.fromisoformat(
                                leave["end"]
                            )
                        )
                    )

        else:
            self.projects = {
                "HOW03": [
                    "CLO"
                ],
                "OW3": [
                    "Drums Unloading",
                    "Sheath Test Rostock"
                ],
                "Oresund2": [
                    "LJB Jointing"
                ]
            }

            self.employees = [
                "APB","GIK","KOB","KOK","LER","PAP","THS"   
            ]
    
    def save_data(self):
        serializable_assignments = {}

        for key, employee in self.assignments.items():
            project, site, assigned_date = key

            serializable_key = (
                f"{project}|{site}|{assigned_date}"
            )

            serializable_assignments[
                serializable_key
            ] = employee

        serializable_site_periods = {}

        for key, period in self.site_periods.items():
            project, site = key

            serializable_key = (
                f"{project}|{site}"
            )

            serializable_site_periods[
                serializable_key
            ] = {
                "start": str(period[0]),
                "end": str(period[1])
            }

        serializable_leaves = {}

        for employee, leaves in self.employee_leaves.items():
            serializable_leaves[employee] = []

            for leave_start, leave_end in leaves:
                serializable_leaves[employee].append({
                    "start": str(leave_start),
                    "end": str(leave_end)
                })
        
        data = {
            "projects": self.projects,
            "employees": self.employees,
            "site_periods": serializable_site_periods,
            "assignments": serializable_assignments,
            "employee_leaves": serializable_leaves
        }

        with open("scheduler_data.json", "w") as f:
            json.dump(data, f, indent=4)
    
    def init_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout()

        # -----------------------------
        # Header Logo
        # -----------------------------
        logo_label = QLabel()

        pixmap = QPixmap("hec_logo.jpg")

        if not pixmap.isNull():
            pixmap = pixmap.scaled(
                300,
                120,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )

            logo_label.setPixmap(pixmap)
        else:
            logo_label.setText("Hellenic Cables")
        
        logo_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(logo_label)

        # -----------------------------
        # Tabs
        # -----------------------------
        tabs = QTabWidget()

        tabs.addTab(
            self.create_dashboard_tab(),
            "📊 Dashboard"
        )

        tabs.addTab(
            self.create_scheduler_tab(),
            "📅 Project Scheduler"
        )

        tabs.addTab(
            self.create_employee_tab(),
            "👷 Employee Availability"
        )

        main_layout.addWidget(tabs)

        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        
    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def load_window_icon(self):
        if os.path.exists("hec_logo.jpg"):
            self.setWindowIcon(QIcon("hec_logo.jpg"))

    # -----------------------------
    # DASHBOARD TAB
    # -----------------------------
    def create_dashboard_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()

        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # Upcoming
        upcoming_box = QGroupBox("Upcoming Operations")
        upcoming_layout = QVBoxLayout()
        self.upcoming_sites_label = QLabel()
        self.upcoming_sites_label.setStyleSheet("font-size:14px;")
        upcoming_layout.addWidget(self.upcoming_sites_label)
        upcoming_box.setLayout(upcoming_layout)

        # Running
        running_box = QGroupBox("Running Operations")
        running_layout = QVBoxLayout()
        self.running_sites_label = QLabel()
        running_layout.addWidget(self.running_sites_label)
        running_box.setLayout(running_layout)

        # Available employees
        available_box = QGroupBox("Available Employees")
        available_layout = QVBoxLayout()
        self.available_employees_label = QLabel()
        available_layout.addWidget(self.available_employees_label)
        available_box.setLayout(available_layout)

        # Leave
        leave_box = QGroupBox("Employees on Leave ✈️🌴🛏️😴💤")
        leave_layout = QVBoxLayout()
        self.leave_employees_label = QLabel()
        leave_layout.addWidget(self.leave_employees_label)
        leave_box.setLayout(leave_layout)
        
        groupbox_style = """
        QGroupBox {
            font-weight: bold;
            font-size: 14px;
            color: black;
            border: 1px solid gray;
            border-radius: 8px;
            margin-top: 10px;
            padding-top: 10px;
        }

        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px 0 5px;
        }
        """
        
        upcoming_box.setStyleSheet(groupbox_style)
        running_box.setStyleSheet(groupbox_style)
        available_box.setStyleSheet(groupbox_style)
        leave_box.setStyleSheet(groupbox_style)

        layout.addWidget(upcoming_box)
        layout.addWidget(running_box)
        layout.addWidget(available_box)
        layout.addWidget(leave_box)
        
        widget.setLayout(layout)

        self.update_dashboard()

        return widget
    
    def update_dashboard(self):
        today = date.today()

        # --------------------------------
        # Upcoming sites
        # --------------------------------
        upcoming_sites = []

        for (project, site), (start, end) in self.site_periods.items():

            if start > today:
                upcoming_sites.append(
                    f"{project} / {site}"
                )

        # --------------------------------
        # Running sites
        # --------------------------------
        running_sites = []

        for (project, site), (start, end) in self.site_periods.items():

            if start <= today <= end:
                running_sites.append(
                    f"{project} / {site}"
                )

        # --------------------------------
        # Employees on leave today
        # --------------------------------
        employees_on_leave = []

        for employee, leaves in self.employee_leaves.items():

            for leave_start, leave_end in leaves:

                if leave_start <= today <= leave_end:
                    employees_on_leave.append(employee)
                    break

        # --------------------------------
        # Assigned employees
        # --------------------------------
        assigned_employees = set()

        for (project, site, assigned_date), employee in self.assignments.items():
            assigned_employees.add(employee)

        # --------------------------------
        # Available employees
        # --------------------------------
        available_employees = []

        for employee in self.employees:
            if (
                employee not in assigned_employees
                and employee not in employees_on_leave
            ):
                available_employees.append(employee)

        # --------------------------------
        # Update labels
        # --------------------------------
        upcoming_text = (
            f"Upcoming Operations/Sites: {len(upcoming_sites)}"
        )

        if upcoming_sites:
            upcoming_text += "\n• " + "\n• ".join(upcoming_sites)

        self.upcoming_sites_label.setText(upcoming_text)


        running_text = (
            f"Running Sites: {len(running_sites)}"
        )

        if running_sites:
            running_text += "\n• " + "\n• ".join(running_sites)

        self.running_sites_label.setText(running_text)


        available_text = (
            f"Available Employees: {len(available_employees)}"
        )

        if available_employees:
            available_text += "\n• " + "\n• ".join(available_employees)

        self.available_employees_label.setText(
            available_text
        )


        leave_text = (
            f"Employees on Leave Today: {len(employees_on_leave)}"
        )

        if employees_on_leave:
            leave_text += "\n• " + "\n• ".join(employees_on_leave)

        self.leave_employees_label.setText(
            leave_text
        )

    # -----------------------------
    # EMPLOYEE TAB
    # -----------------------------
    def create_employee_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()

        layout.addWidget(QLabel("Employee Availability / Assignments"))

        # employee list
        self.employee_assignment_view = QListWidget()
        layout.addWidget(self.employee_assignment_view)
        
        leave_controls = QHBoxLayout()

        self.leave_start = QDateEdit()
        self.leave_start.setCalendarPopup(True)
        self.leave_start.setDate(QDate.currentDate())

        self.leave_end = QDateEdit()
        self.leave_end.setCalendarPopup(True)
        self.leave_end.setDate(QDate.currentDate())

        self.leave_employee_dropdown = QComboBox()
        self.leave_employee_dropdown.addItems(self.employees)

        self.add_leave_btn = QPushButton("Add Leave")
        self.add_leave_btn.clicked.connect(self.add_employee_leave)

        leave_controls.addWidget(QLabel("Leave Start:"))
        leave_controls.addWidget(self.leave_start)

        leave_controls.addWidget(QLabel("Leave End:"))
        leave_controls.addWidget(self.leave_end)

        leave_controls.addWidget(QLabel("Employee:"))
        leave_controls.addWidget(self.leave_employee_dropdown)

        leave_controls.addWidget(self.add_leave_btn)

        layout.addLayout(leave_controls)

        # buttons
        employee_buttons = QHBoxLayout()

        self.add_employee_btn = QPushButton("Add Employee")
        self.add_employee_btn.clicked.connect(self.add_employee)

        self.remove_employee_btn = QPushButton("Remove Employee")
        self.remove_employee_btn.clicked.connect(self.remove_employee)

        employee_buttons.addWidget(self.add_employee_btn)
        employee_buttons.addWidget(self.remove_employee_btn)

        layout.addLayout(employee_buttons)

        widget.setLayout(layout)

        self.update_employee_tab()

        return widget
    
    def add_employee(self):
        employee_name, ok = QInputDialog.getText(
            self,
            "Add Employee",
            "Employee name:"
        )

        if ok and employee_name:

            if employee_name in self.employees:
                QMessageBox.warning(
                    self,
                    "Error",
                    "Employee already exists"
                )
                return

            self.employees.append(employee_name)

            # update dropdown
            self.employee_dropdown.clear()
            self.employee_dropdown.addItems(self.employees)

            self.update_employee_tab()
            
        self.save_data()
            
    def remove_employee(self):
        employee_name, ok = QInputDialog.getItem(
            self,
            "Remove Employee",
            "Select employee:",
            self.employees,
            0,
            False
        )

        if not ok or not employee_name:
            return

        # remove assignments first
        assignments_to_remove = []

        for key, assigned_employee in self.assignments.items():
            if assigned_employee == employee_name:
                assignments_to_remove.append(key)

        for key in assignments_to_remove:
            del self.assignments[key]

        # remove employee
        self.employees.remove(employee_name)

        # update dropdown
        self.employee_dropdown.clear()
        self.employee_dropdown.addItems(self.employees)

        # rebuild timeline colors
        self.rebuild_timeline()

        self.update_employee_tab()
        
        self.save_data()
    
    def update_employee_tab(self):
        self.employee_assignment_view.clear()

        font = self.employee_assignment_view.font()
        bold_font = font
        bold_font.setBold(True)

        for employee in self.employees:

            employee_assignments = []

            # collect assignments for current employee
            for (
                project_name,
                site_name,
                assigned_date
            ), assigned_employee in self.assignments.items():

                if assigned_employee == employee:
                    employee_assignments.append(
                        (
                            project_name,
                            site_name,
                            assigned_date
                        )
                    )

            # sort by date
            employee_assignments.sort(
                key=lambda x: x[2]
            )

            # ------------------------
            # employee header
            # ------------------------
            employee_item = QListWidgetItem(employee)
            employee_item.setFont(bold_font)
            self.employee_assignment_view.addItem(employee_item)

            if not employee_assignments:
                self.employee_assignment_view.addItem(
                    "   Available"
                )
                self.employee_assignment_view.addItem("")
                continue

            # ------------------------
            # group consecutive dates
            # ------------------------
            grouped = []

            current_project = employee_assignments[0][0]
            current_site = employee_assignments[0][1]
            start_date = employee_assignments[0][2]
            prev_date = employee_assignments[0][2]

            for assignment in employee_assignments[1:]:
                project, site, current_date = assignment

                same_site = (
                    project == current_project
                    and site == current_site
                )

                consecutive = (
                    current_date - prev_date
                ).days == 1

                if same_site and consecutive:
                    prev_date = current_date

                else:
                    grouped.append(
                        (
                            current_project,
                            current_site,
                            start_date,
                            prev_date
                        )
                    )

                    current_project = project
                    current_site = site
                    start_date = current_date
                    prev_date = current_date

            # final group
            grouped.append(
                (
                    current_project,
                    current_site,
                    start_date,
                    prev_date
                )
            )

            # ------------------------
            # display grouped ranges
            # ------------------------
            for project, site, start, end in grouped:

                if start == end:
                    text = (
                        f"   • {project} / {site} "
                        f"→ {start}"
                    )
                else:
                    text = (
                        f"   • {project} / {site} "
                        f"→ {start} until {end}"
                    )

                self.employee_assignment_view.addItem(text)

            # ------------------------
            # show leave periods
            # ------------------------
            if employee in self.employee_leaves:
                for leave_start, leave_end in self.employee_leaves[employee]:

                    if leave_start == leave_end:
                        leave_text = (
                            f"   ✈️🌴🛏️😴💤 Leave → {leave_start}"
                        )
                    else:
                        leave_text = (
                            f"   ✈️🌴🛏️😴💤 Leave → "
                            f"{leave_start} until {leave_end}"
                        )

                    self.employee_assignment_view.addItem(
                        leave_text
                    )
            
            # spacing between employees
            self.employee_assignment_view.addItem("")

    def add_employee_leave(self):
        employee = self.leave_employee_dropdown.currentText()

        start = self.leave_start.date().toPython()
        end = self.leave_end.date().toPython()

        if start > end:
            QMessageBox.warning(
                self,
                "Error",
                "Invalid leave dates"
            )
            return

        if employee not in self.employee_leaves:
            self.employee_leaves[employee] = []

        self.employee_leaves[employee].append(
            (start, end)
        )

        QMessageBox.information(
            self,
            "Success",
            f"{employee} leave added"
        )

        self.update_employee_tab()
        self.update_dashboard()
        self.save_data()
            
    # -----------------------------
    # SCHEDULER TAB
    # -----------------------------
    def create_scheduler_tab(self):
        widget = QWidget()
        main_layout = QVBoxLayout()

        # Top filters
        filter_layout = QHBoxLayout()

        self.project_filter = QComboBox()
        self.project_filter.addItem("All Projects")
        self.project_filter.addItems(self.projects.keys())
        
        self.project_filter.currentTextChanged.connect(
            self.apply_filters
        )

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search site...")
        
        self.search_box.textChanged.connect(
            self.apply_filters
        )

        filter_layout.addWidget(QLabel("Project Filter:"))
        filter_layout.addWidget(self.project_filter)
        filter_layout.addWidget(self.search_box)

        main_layout.addLayout(filter_layout)

        # Split view
        splitter = QSplitter(Qt.Horizontal)

        # LEFT PANEL -> hierarchy
        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Projects / Sites")
        self.populate_tree()

        # right-click menu
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.open_tree_menu)

        # RIGHT PANEL -> timeline
        self.timeline = QTableWidget()
        self.timeline.setAlternatingRowColors(True)
        self.setup_timeline()
        self.timeline.setSelectionMode(QTableWidget.ExtendedSelection)
        self.timeline.setSelectionBehavior(QTableWidget.SelectItems)
        self.restore_timeline_state()
        self.timeline.cellDoubleClicked.connect(self.assign_employee)

        splitter.addWidget(self.tree)
        splitter.addWidget(self.timeline)
        splitter.setSizes([200, 1400])

        main_layout.addWidget(splitter)

        # Controls
        controls = QHBoxLayout()

        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate())

        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())

        self.create_period_btn = QPushButton("Create Work Period")
        self.create_period_btn.clicked.connect(self.create_work_period)

        controls.addWidget(QLabel("First Site Day:"))
        controls.addWidget(self.start_date)
        controls.addWidget(QLabel("Last Site Day:"))
        controls.addWidget(self.end_date)
        controls.addWidget(self.create_period_btn)

        main_layout.addLayout(controls)
        
        # Assign selected timeline cells
        self.assign_selected_btn = QPushButton("Assign Employee to Selected Cells")
        self.assign_selected_btn.clicked.connect(self.assign_selected_cells)

        main_layout.addWidget(self.assign_selected_btn)
        
        self.remove_assignment_btn = QPushButton("Remove Employee Assignment")
        self.remove_assignment_btn.clicked.connect(self.remove_selected_assignment)

        main_layout.addWidget(self.remove_assignment_btn)
        
        # Employee assignment controls
        assignment_controls = QHBoxLayout()

        self.assign_start_date = QDateEdit()
        self.assign_start_date.setCalendarPopup(True)
        self.assign_start_date.setDate(QDate.currentDate())

        self.assign_end_date = QDateEdit()
        self.assign_end_date.setCalendarPopup(True)
        self.assign_end_date.setDate(QDate.currentDate())

        self.employee_dropdown = QComboBox()
        self.employee_dropdown.addItems(self.employees)

        self.assign_range_btn = QPushButton("Assign Employee Range")
        self.assign_range_btn.clicked.connect(self.assign_employee_range)

        assignment_controls.addWidget(QLabel("Assign Employee Start:"))
        assignment_controls.addWidget(self.assign_start_date)

        assignment_controls.addWidget(QLabel("Assign Employee End:"))
        assignment_controls.addWidget(self.assign_end_date)

        assignment_controls.addWidget(QLabel("Employee:"))
        assignment_controls.addWidget(self.employee_dropdown)

        assignment_controls.addWidget(self.assign_range_btn)

        main_layout.addLayout(assignment_controls)

        widget.setLayout(main_layout)

        return widget

    # -----------------------------
    def open_tree_menu(self, position):
        item = self.tree.itemAt(position)

        menu = QMenu()
        add_project_action = menu.addAction("Add Project")
        add_site_action = menu.addAction("Add Site")
        remove_action = menu.addAction("Remove")

        action = menu.exec(self.tree.viewport().mapToGlobal(position))

        if action == add_project_action:
            self.add_project()

        elif action == add_site_action:
            if item and not item.parent():
                self.add_site(item)
            else:
                QMessageBox.warning(self, "Error", "Select a project first")

        elif action == remove_action and item:
            self.remove_item(item)
            
    def apply_filters(self):
        selected_project = self.project_filter.currentText()
        search_text = self.search_box.text().lower()

        for row in range(self.timeline.rowCount()):
            header_item = self.timeline.verticalHeaderItem(row)

            if not header_item:
                continue

            row_label = header_item.text()
            row_label_lower = row_label.lower()

            # project filter check
            project_match = (
                selected_project == "All Projects"
                or row_label.startswith(selected_project)
            )

            # search filter check
            search_match = (
                search_text == ""
                or search_text in row_label_lower
            )

            # show row only if BOTH conditions pass
            if project_match and search_match:
                self.timeline.setRowHidden(row, False)
            else:
                self.timeline.setRowHidden(row, True)

    def add_project(self):
        project_name, ok = QInputDialog.getText(self, "Add Project", "Project name:")

        if ok and project_name:
            if project_name in self.projects:
                QMessageBox.warning(self, "Error", "Project already exists")
                return

            self.projects[project_name] = []
            project_item = QTreeWidgetItem([project_name])
            self.tree.addTopLevelItem(project_item)
            
        self.save_data()

    def add_site(self, project_item):
        project_name = project_item.text(0)
        current_site_key, ok = QInputDialog.getText(self, "Add Site", "Site name:")

        if ok and current_site_key:
            if current_site_key in self.site_to_row:
                QMessageBox.warning(self, "Error", "Site already exists")
                return

            self.projects[project_name].append(current_site_key)
            site_item = QTreeWidgetItem([current_site_key])
            project_item.addChild(site_item)

            new_row = self.timeline.rowCount()
            self.timeline.insertRow(new_row)
            self.timeline.setVerticalHeaderItem(new_row, QTableWidgetItem(current_site_key))
            self.site_to_row[current_site_key] = new_row
            
        self.save_data()

    def remove_item(self, item):
        if item.parent():
            current_site_key = item.text(0).split(" (")[0]
            project_name = item.parent().text(0)

            if current_site_key in self.projects[project_name]:
                self.projects[project_name].remove(current_site_key)

            item.parent().removeChild(item)

        else:
            project_name = item.text(0)
            del self.projects[project_name]
            index = self.tree.indexOfTopLevelItem(item)
            self.tree.takeTopLevelItem(index)

        # rebuild timeline after every delete
        self.rebuild_timeline()
        
        self.save_data()

    def rebuild_timeline(self):
        self.timeline.clearContents()
        self.setup_timeline()

        self.restore_site_periods()
        self.restore_assignments()
        
    def restore_site_periods(self):
        for (project_name, site_name), period in self.site_periods.items():

            start, end = period

            row = self.site_to_row[(project_name, site_name)]

            for col, d in enumerate(self.dates):

                if d < start:
                    color = QColor("lightgray")

                elif start <= d <= end:
                    color = QColor("yellow")

                else:
                    color = QColor("lightgray")

                cell = self.timeline.item(row, col)

                if not cell:
                    cell = QTableWidgetItem("")
                    self.timeline.setItem(row, col, cell)

                cell.setBackground(color)

    def restore_assignments(self):
        for (project_name, site_name, assigned_date), employee in self.assignments.items():

            row = self.site_to_row[(project_name, site_name)]

            if assigned_date not in self.dates:
                continue

            col = self.dates.index(assigned_date)

            cell = self.timeline.item(row, col)

            if not cell:
                cell = QTableWidgetItem("")
                self.timeline.setItem(row, col, cell)

            cell.setText(employee)
            cell.setTextAlignment(Qt.AlignCenter)
            cell.setBackground(QColor("#90EE90"))
    
    def populate_tree(self):
        for project, sites in self.projects.items():
            project_item = QTreeWidgetItem([project])
            self.tree.addTopLevelItem(project_item)

            for site in sites:
                site_item = QTreeWidgetItem([site])
                project_item.addChild(site_item)

    # -----------------------------
    def setup_timeline(self):
        today = date.today()

        # Monday = 0
        start = today - timedelta(days=today.weekday())

        self.dates = [
            start + timedelta(days=i)
            for i in range(180)
        ]

        active_sites = list(self.site_periods.keys())

        self.timeline.setColumnCount(len(self.dates))
        self.timeline.setRowCount(len(active_sites))
        
        self.timeline.verticalHeader().setDefaultSectionSize(45)

        headers = [d.strftime("%d/%m") for d in self.dates]
        self.timeline.setHorizontalHeaderLabels(headers)
        for col in range(len(self.dates)):
            self.timeline.setColumnWidth(col, 55)

        self.site_to_row = {}

        for row, (project, site) in enumerate(active_sites):
            label = f"{project} - {site}"

            self.site_to_row[(project, site)] = row

            self.timeline.setVerticalHeaderItem(
                row,
                QTableWidgetItem(label)
            )

            start_dt, end_dt = self.site_periods[(project, site)]

            for col, d in enumerate(self.dates):

                cell = QTableWidgetItem("")

                if d < start_dt:
                    cell.setBackground(QColor("#f7f9fc"))

                elif start_dt <= d <= end_dt:
                    cell.setBackground(QColor("#ffe599"))

                else:
                    cell.setBackground(QColor("#f7f9fc"))

                # restore employee assignment if exists
                assignment_key = ((project, site), d)

                if assignment_key in self.assignments:
                    employee = self.assignments[assignment_key]

                    cell.setText(employee)
                    cell.setTextAlignment(Qt.AlignCenter)
                    cell.setBackground(QColor("#90EE90"))

                self.timeline.setItem(row, col, cell)

    def restore_timeline_state(self):
        # -------------------------
        # restore active periods
        # -------------------------
        for (project_name, site_name), (start, end) in self.site_periods.items():

            site_key = (project_name, site_name)

            if site_key not in self.site_to_row:
                continue

            row = self.site_to_row[site_key]

            for col, d in enumerate(self.dates):

                if start <= d <= end:
                    cell = self.timeline.item(row, col)

                    if cell is None:
                        cell = QTableWidgetItem("")
                        self.timeline.setItem(row, col, cell)

                    cell.setBackground(QColor("yellow"))

        # -------------------------
        # restore employee assignments
        # -------------------------
        for (
            project_name,
            site_name,
            assigned_date
        ), employee in self.assignments.items():

            site_key = (project_name, site_name)

            if site_key not in self.site_to_row:
                continue

            row = self.site_to_row[site_key]

            if assigned_date not in self.dates:
                continue

            col = self.dates.index(assigned_date)

            cell = self.timeline.item(row, col)

            if cell is None:
                cell = QTableWidgetItem("")
                self.timeline.setItem(row, col, cell)

            cell.setText(employee)
            cell.setTextAlignment(Qt.AlignCenter)
            cell.setBackground(QColor("#B7E1CD"))
                
    # -----------------------------
    def create_work_period(self):
        item = self.tree.currentItem()

        if not item or not item.parent():
            QMessageBox.warning(
                self,
                "Error",
                "Select a site first"
            )
            return
        
        item = self.tree.currentItem()
        
        # καθαρό site name (χωρίς τις ημερομηνίες)
        site_name = item.text(0).split(" (")[0]

        # parent project
        project_name = item.parent().text(0)

        # unique identifier
        site_key = (project_name, site_name)

        start = self.start_date.date().toPython()
        end = self.end_date.date().toPython()

        if start > end:
            QMessageBox.warning(
                self,
                "Error",
                "Start date cannot be after end date"
            )
            return

        # αποθήκευση active period
        self.site_periods[site_key] = (start, end)

        # update tree label
        self.update_site_label(project_name, site_name)

        # rebuild timeline
        self.rebuild_timeline()

        QMessageBox.information(
            self,
            "Success",
            f"{project_name} - {site_name} activated "
            f"from {start} to {end}"
        )
        
        self.update_dashboard()
        self.save_data()
        
    def assign_selected_cells(self):
        selected_items = self.timeline.selectedItems()

        if not selected_items:
            QMessageBox.warning(
                self,
                "Error",
                "Select timeline cells first"
            )
            return

        employee, ok = QInputDialog.getItem(
            self,
            "Assign Employee",
            "Select Employee:",
            self.employees,
            0,
            False
        )

        if not ok:
            return

        for item in selected_items:
            row = item.row()
            col = item.column()

            if item.background().color() != QColor("yellow"):
                continue

            selected_date = self.dates[col]

            current_site_key = None

            for site_key, r in self.site_to_row.items():
                if r == row:
                    current_site_key = site_key
                    break

            if current_site_key is None:
                continue

            project_name, site_name = current_site_key

            # --------------------------
            # LEAVE CHECK
            # --------------------------
            if employee in self.employee_leaves:
                on_leave = False

                for leave_start, leave_end in self.employee_leaves[employee]:
                    if leave_start <= selected_date <= leave_end:
                        on_leave = True
                        break

                if on_leave:
                    QMessageBox.warning(
                        self,
                        "Leave Conflict",
                        f"{employee} is on leave on {selected_date}"
                    )
                    continue

            # --------------------------
            # DOUBLE BOOKING CHECK
            # --------------------------
            conflict_found = False

            for (
                project,
                site,
                assigned_date
            ), assigned_employee in self.assignments.items():

                if (
                    assigned_employee == employee
                    and assigned_date == selected_date
                    and (project, site) != current_site_key
                ):
                    QMessageBox.warning(
                        self,
                        "Conflict",
                        f"{employee} already assigned to "
                        f"{project}/{site} on {selected_date}"
                    )
                    conflict_found = True
                    break

            if conflict_found:
                continue

            item.setText(employee)
            item.setBackground(QColor("#B7E1CD"))

            self.assignments[
                (project_name, site_name, selected_date)
            ] = employee

        self.update_employee_tab()
        self.update_dashboard()
        self.save_data()
        
    def remove_selected_assignment(self):
        selected_items = self.timeline.selectedItems()

        if not selected_items:
            QMessageBox.warning(
                self,
                "Error",
                "Select assigned cells first"
            )
            return

        for item in selected_items:
            row = item.row()
            col = item.column()

            current_color = item.background().color()

            # remove only assigned cells
            if current_color != QColor("#90EE90"):
                continue

            selected_date = self.dates[col]

            current_site_key = None

            for site_key, r in self.site_to_row.items():
                if r == row:
                    current_site_key = site_key
                    break

            if current_site_key is None:
                continue

            # unpack tuple
            project_name, site_name = current_site_key

            # remove text
            item.setText("")

            # revert back to active period
            item.setBackground(QColor("yellow"))

            # remove assignment from dictionary
            assignment_key = (
                project_name,
                site_name,
                selected_date
            )

            if assignment_key in self.assignments:
                del self.assignments[assignment_key]

        self.update_employee_tab()
        self.update_dashboard()
        self.save_data()
    
    def assign_employee_range(self):
        item = self.tree.currentItem()

        if not item or not item.parent():
            QMessageBox.warning(
                self,
                "Error",
                "Select a site first"
            )
            return

        site_name = item.text(0).split(" (")[0]
        project_name = item.parent().text(0)

        current_site_key = (project_name, site_name)

        assign_start = self.assign_start_date.date().toPython()
        assign_end = self.assign_end_date.date().toPython()

        employee = self.employee_dropdown.currentText()

        row = self.site_to_row[current_site_key]

        for col, d in enumerate(self.dates):

            if assign_start <= d <= assign_end:

                cell = self.timeline.item(row, col)

                if not cell:
                    continue

                if cell.background().color() != QColor("yellow"):
                    continue

                # --------------------------
                # LEAVE CHECK
                # --------------------------
                if employee in self.employee_leaves:
                    on_leave = False

                    for leave_start, leave_end in self.employee_leaves[employee]:
                        if leave_start <= d <= leave_end:
                            on_leave = True
                            break

                    if on_leave:
                        QMessageBox.warning(
                            self,
                            "Leave Conflict",
                            f"{employee} is on leave on {d}"
                        )
                        continue

                # --------------------------
                # DOUBLE BOOKING CHECK
                # --------------------------
                conflict_found = False

                for (
                    project,
                    site,
                    assigned_date
                ), assigned_employee in self.assignments.items():

                    if (
                        assigned_employee == employee
                        and assigned_date == d
                        and (project, site) != current_site_key
                    ):
                        QMessageBox.warning(
                            self,
                            "Conflict",
                            f"{employee} already assigned to "
                            f"{project}/{site} on {d}"
                        )
                        conflict_found = True
                        break

                if conflict_found:
                    continue

                cell.setText(employee)
                cell.setTextAlignment(Qt.AlignCenter)
                cell.setBackground(QColor("#B7E1CD"))

                self.assignments[
                    (project_name, site_name, d)
                ] = employee

        self.update_employee_tab()
        self.update_dashboard()
        self.save_data()

    def update_site_label(self, project_name, site_name):
        site_key = (project_name, site_name)

        if site_key not in self.site_periods:
            return

        start, end = self.site_periods[site_key]

        formatted_start = start.strftime("%d/%m/%Y")
        formatted_end = end.strftime("%d/%m/%Y")

        for i in range(self.tree.topLevelItemCount()):
            project_item = self.tree.topLevelItem(i)

            if project_item.text(0) == project_name:

                for j in range(project_item.childCount()):
                    site_item = project_item.child(j)

                    clean_site_name = site_item.text(0).split(" (")[0]

                    if clean_site_name == site_name:
                        site_item.setText(
                            0,
                            f"{site_name} ({formatted_start} - {formatted_end})"
                        )
                        return
        
    # -----------------------------
    def assign_employee(self, row, column):
        item = self.timeline.item(row, column)

        if item is None:
            return

        if item.background().color() != QColor("yellow"):
            QMessageBox.warning(
                self,
                "Error",
                "You can only assign employees in active periods."
            )
            return

        employee, ok = QInputDialog.getItem(
            self,
            "Assign Employee",
            "Select Employee:",
            self.employees,
            0,
            False
        )

        if not ok:
            return

        selected_date = self.dates[column]

        current_site_key = None

        for site_key, r in self.site_to_row.items():
            if r == row:
                current_site_key = site_key
                break

        if current_site_key is None:
            return

        project_name, site_name = current_site_key

        # --------------------------
        # LEAVE CHECK
        # --------------------------
        if employee in self.employee_leaves:
            for leave_start, leave_end in self.employee_leaves[employee]:
                if leave_start <= selected_date <= leave_end:
                    QMessageBox.warning(
                        self,
                        "Leave Conflict",
                        f"{employee} is on leave on {selected_date}"
                    )
                    return

        # --------------------------
        # DOUBLE BOOKING CHECK
        # --------------------------
        for (
            project,
            site,
            assigned_date
        ), assigned_employee in self.assignments.items():

            if (
                assigned_employee == employee
                and assigned_date == selected_date
                and (project, site) != current_site_key
            ):
                QMessageBox.warning(
                    self,
                    "Conflict Detected",
                    f"{employee} is already assigned to "
                    f"{project}/{site} on {selected_date}"
                )
                return

        item.setText(employee)
        item.setBackground(QColor("#B7E1CD"))

        self.assignments[
            (project_name, site_name, selected_date)
        ] = employee

        self.update_employee_tab()
        self.save_data()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SchedulerApp()
    window.show()
    sys.exit(app.exec())


# In[ ]:




