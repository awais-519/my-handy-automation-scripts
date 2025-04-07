from SalarySlipsManager.main import Automation as SalarySlipsAutomation

def main():
    try:
        print("\n\nWhich automation do you want to run?\n\n1. Salary Slips Manager\n\n")
        option = input("Enter your choice: ").strip()

        if option == "1":
            automation = SalarySlipsAutomation()
            automation.run()
        else:
            print("You had one job to do!")
            exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        exit(1)
    finally:
        exit(0)

if __name__ == "__main__":
    main()