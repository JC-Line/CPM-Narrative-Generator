# CPM-Narrative-Generator

## Purpose
Quickly create a word document filled with Narrative information required for monthly FDOT CPM schedule Submittals

## Usage
- Intended usage is through using PyInstaller to package the program as an executable
- Can be ran directly through code interpretation

- Running the application will prompt the user for the following information
    - Comparison Schedule: .XLSX formatted [P6 Export](#p6-export-formatting)
    - Comparison Data Date: Data date of the comparison shcedule, should be older than current
    - Current Schedule: .XLSX formatted [P6 Export](#p6-export-formatting)
    - Current Schedule Data Date: Data date of the current schedule, should be newer than comparison
    - Word Template: .DOCX formatted file to be used as a template for the generated narrative
        - Template will not be modified but needs to contain the desired formatting for the application to use
    
### P6 Export Formatting
Export must include both Activity and Relationship Data
    - Activity Data
        - task_code	
        - status_code
        - wbs_id
        - wbs_name
        - task_name
        - critical_flag
        - start_date
        - end_date
        - act_start_date
        - act_end_date
        - target_drtn_hr_cnt
        - act_drtn_hr_cnt
        - remain_drtn_hr_cnt
        - total_float_hr_cnt
        - float_path
        - resource_list
    - Relationship Data
        - pred_task_id
        - task_id
        - pred_type
        - PREDTASK__status_code
        - TASK__status_code
        - pred_proj_id
        - proj_id
        - predtask__projwbs__wbs_full_name
        - task__projwbs__wbs_full_name
        - predtask__task_name
        - task__task_name
        - lag_hr_cnt
        - predtask__rsrc_id
        - task__rsrc_id
        - comments

## Updates Needed
- Update table formatting
