import pandas as pd
import os

def read_file(file, length = None):
    if length:
        df = pd.read_excel(file, sheet_name="工作表1",
                        usecols=["工作站(製程)", "工站代號", "工站簡稱", "寬放工時(min)"],
                        nrows = length
                        )
    else:
        df = pd.read_excel(file, sheet_name="工作表1",
                        usecols=["工作站(製程)", "工站代號", "工站簡稱", "寬放工時(min)"]
                        )
    return df

def find_min_idx(ptime, best_time):
    dif_list = []
    current = 0

    for i in range(len(ptime)):
        current += ptime[i]
        dif = abs(current - best_time)
        dif_list.append(dif)

    min_idx = dif_list.index(min(dif_list))
    
    return min_idx

def assignment_alg(dic, num_of_workers)-> None:
    ptime = dic["寬放工時(min)"]
    cycletime = sum(ptime)
    best_time = cycletime / num_of_workers
    residual = len(ptime)
    
    start_idx = 0
    assignments = []
    new_process_mapping = {}
    new_process_number = 1

    while start_idx < residual:
        min_idx = find_min_idx(ptime[start_idx:], best_time)
        end_idx = start_idx + min_idx + 1
        if end_idx > residual:
            end_idx = residual
        assignments.append((start_idx, end_idx))

        new_process_mapping[new_process_number] = {
            "Original Process Numbers": dic["工站代號"][start_idx:end_idx],
            "Detailed Info": {
                "工作站(製程)": dic["工作站(製程)"][start_idx:end_idx],
                "工站簡稱": dic["工站簡稱"][start_idx:end_idx],
                "寬放工時(min)": dic["寬放工時(min)"][start_idx:end_idx],
                "新工站工時": sum(ptime[start_idx: end_idx])
            }
        }
        
        start_idx = end_idx
        new_process_number += 1

    # Convert the mapping to a DataFrame and save to an Excel file
    output_data = []

    for new_proc, details in new_process_mapping.items():
        for i in range(len(details['Original Process Numbers'])):
            output_data.append({
                "新工站代號": new_proc,
                "原工站代號": details['Original Process Numbers'][i],
                "工作站(製程)": details["Detailed Info"]["工作站(製程)"][i],
                "工站簡稱": details["Detailed Info"]["工站簡稱"][i],
                "寬放工時(min)": details["Detailed Info"]["寬放工時(min)"][i],
                "新工站工時": details["Detailed Info"]["新工站工時"]
            })
    
    output_df = pd.DataFrame(output_data)
    output_path = r"C:\Users\USER\Desktop\Python\MTNet\output.xlsx"
    
    if not os.path.exists(output_path):
        output_df.to_excel(output_path, index=False)

    print("Assignments:")
    
    for i, (start, end) in enumerate(assignments):
        print(f"Workstation {i+1} assigned to processes {dic['工作站(製程)'][start]} to {dic['工作站(製程)'][end-1]}")
    return output_df

def main():
    seq_length = 33
    num_of_workers = 7
    file_path = r"C:\Users\USER\Desktop\Python\MTNet\Input.xlsx"
    df = read_file(file_path, seq_length)
    dic = df.to_dict(orient="list")
    assignment_alg(dic, num_of_workers)



if __name__ == "__main__":
    main()
