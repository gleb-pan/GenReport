import argparse

# Define your existing functions
def gen_device_report():
    # Your existing gen_device_report implementation
    print("Generating device report...")

def other_function():
    # Another function example
    print("Running another function...")

# Define a function mapper dictionary
function_mapper = {
    'gen_device_report': gen_device_report,
    'other_function': other_function,
}

# Parse command-line arguments
def main():
    parser = argparse.ArgumentParser(description='Run a function based on command-line argument.')
    parser.add_argument('function_name', choices=function_mapper.keys(), help='Name of the function to run.')
    args = parser.parse_args()

    # Call the selected function
    selected_function = function_mapper[args.function_name]
    selected_function()

if __name__ == '__main__':
    main()