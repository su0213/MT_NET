import random
import math
import numpy as np
import pandas as pd
from WS_Consoildation import read_file, assignment_alg 

# 計算總適應度
def calculate_total_adaptability(assignment, adaptability_matrix):
    total_adaptability = 0
    for person, machine in enumerate(assignment):
        total_adaptability += adaptability_matrix[person][machine]
    return total_adaptability

# 生成初始分配
def initial_assignment(n_people, n_machines):
    assignment = list(range(n_machines))
    random.shuffle(assignment)
    return assignment

# 模擬退火算法
def simulated_annealing(adaptability_matrix, initial_temp, cooling_rate, max_iterations):
    n_people = len(adaptability_matrix)
    n_machines = len(adaptability_matrix[0])
    
    current_assignment = initial_assignment(n_people, n_machines)
    current_adaptability = calculate_total_adaptability(current_assignment, adaptability_matrix)
    temperature = initial_temp

    iteration = 0
    while temperature > 1 and iteration < max_iterations:
        new_assignment = current_assignment[:]
        i, j = random.sample(range(n_machines), 2)  # 隨機交換兩個崗位
        new_assignment[i], new_assignment[j] = new_assignment[j], new_assignment[i]
        
        new_adaptability = calculate_total_adaptability(new_assignment, adaptability_matrix)

        if new_adaptability > current_adaptability or random.uniform(0, 1) < math.exp((new_adaptability - current_adaptability) / temperature):
            current_assignment = new_assignment
            current_adaptability = new_adaptability

        temperature *= cooling_rate
        iteration += 1

    return current_assignment, current_adaptability

def Get_input_as_dict(seq_length, file_path, num_of_workers):
    file = read_file(file_path, seq_length)
    dic = file.to_dict(orient="list")
    input = assignment_alg(dic, num_of_workers)
    return input

def Get_employee_skill(path):
    df = pd.read_excel(path, sheet_name="工作表1", usecols="B, C, D, E, F, G, H, I, J, K")
    df = df.T
    dic = df.to_dict(orient="list")
    return dic

def organize_station_data(input_dict):
    stations = {}
    for idx, (new_station, old_station, process, _, _) in enumerate(zip(
        input_dict['新工站代號'],
        input_dict['原工站代號'],
        input_dict['工作站(製程)'],
        input_dict['寬放工時(min)'],
        input_dict['新工站工時']
    )):
        if new_station not in stations:
            stations[new_station] = {"工站代號": new_station, "所需製程": [], "包含工作站": []}
        stations[new_station]["所需製程"].append(old_station)
        stations[new_station]["包含工作站"].append(process)
    return stations

def organize_employee_skills(skill_info_dic):
    people = {}
    for idx, info in skill_info_dic.items():
        role, age, exp, B, A, P, J, S, D, T = info
        people[idx] = [role, age, exp, B, A, P, J, S, D, T]
    return people

def build_adaptability_matrix(stations, people):
    # 定義技能索引對應
    skill_indices = {
        "B": 3,
        "A": 4,
        "P": 5,
        "J": 6,
        "S": 7,
        "D": 8,
        "T": 9 
    }

    # 提取每個工站所需的技能索引
    station_skills = {}
    for station_id, info in stations.items():
        skills = list(set([skill_indices[process] for process in info["所需製程"]]))
        station_skills[station_id] = skills

    n_people = len(people)
    n_stations = len(stations)
    adaptability_matrix = np.zeros((n_people, n_stations), dtype=int)

    for person_id, person_info in people.items():
        person_skills = person_info[3:]  # 提取人員技能部分
        for station_id, skills_needed in station_skills.items():
            score = sum(person_skills[skill_index - 3] for skill_index in skills_needed)  # 計算匹配得分
            adaptability_matrix[person_id][station_id - 1] = score  # 確保station_id從0開始
    return adaptability_matrix

def select_top_k_employees(adaptability_matrix, k):
    total_adaptability_scores = adaptability_matrix.sum(axis=1)
    top_k_indices = np.argsort(total_adaptability_scores)[-k:]  # 選擇適應度最高的k個員工
    return top_k_indices

def main():
    seq_length = 33
    num_of_workstations = 7
    Consoildation_algo_file_path = r"C:\Users\USER\Desktop\Python\MTNet\Input.xlsx"
    Employee_skill_path = r"C:\Users\USER\Desktop\Python\MTNet\Employee_skill_info.xlsx"
    
    input = Get_input_as_dict(seq_length, Consoildation_algo_file_path, num_of_workstations)
    skill_info_dic = Get_employee_skill(Employee_skill_path)
    
    station = organize_station_data(input)
    people = organize_employee_skills(skill_info_dic)
    adaptability_matrix = build_adaptability_matrix(station, people)

    top_k_indices = select_top_k_employees(adaptability_matrix, num_of_workstations)
    top_k_adaptability_matrix = adaptability_matrix[top_k_indices]

    iterations = 1000
    initial_temp = 100
    cooling_rate = 0.995
    max_iterations = 100000 # 設置最大迭代次數

    # 運行模擬退火算法
    for _ in range(iterations):
        optimal_assignment, optimal_adaptability = simulated_annealing(top_k_adaptability_matrix, initial_temp, cooling_rate, max_iterations)
        if optimal_adaptability == 12:
            break
    
    for i, person_id in enumerate(top_k_indices):
        print(f"Person {person_id + 1}: Machine {optimal_assignment[i] + 1}")

    print(f"Fitness: {optimal_adaptability}")
    
if __name__ == "__main__":
    main()
