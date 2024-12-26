import pandas
import numpy
import matplotlib.pyplot as plt
import sympy
import math
import os

from openpyxl.workbook import Workbook


#Total cost calculation
def cost(Q, D, S, H):
    return (D/Q)*S + (Q/2)*H

#EOQ Calculation
def eoq(D, S, H):
    return math.sqrt(2*D*S/H)

def get_positive_n(prompt):
    user_input = float(input(prompt))
    if user_input <= 0:
        raise ValueError("Input must be positive. Try again.")
    else:
        return user_input


#User data input
def main():

    path = "C:/Users/User/PycharmProjects/PythonProject/Inventory_manager.xlsx"
    while True:
        try:
            Q = get_positive_n("Enter Order Quantity:")
            D = get_positive_n("Enter Total Demand:")
            S = get_positive_n("Enter Cost/order:")
            H = get_positive_n("Enter Holding Cost:")
            table = pandas.DataFrame({"Order Quantity": [Q],
                                "Total Demand": [D],
                                "Order Cost": [S],
                                "Holding Cost": [H],
                                "Total Cost": [cost(Q, D, S, H)],
                                "EOQ": [eoq(D, S, H)]})

            print(table)
            if os.path.exists(path):
                existing_data = pandas.read_excel(path)
                updated_data = pandas.concat([existing_data, table])
            else:
                updated_data = table

            updated_data.to_excel(path, index=False)


        except ValueError as e:
            print("Please enter a number")
        except FileNotFoundError as e:
            print("File not found")
        except ZeroDivisionError as e:
            print("Holding cost cannot be zero")
        else:
            break
        finally:
            pass

if __name__ == "__main__":
    main()
