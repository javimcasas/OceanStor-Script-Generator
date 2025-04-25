import os
import unittest
import datetime
from importlib import import_module
from inspect import getmembers, isfunction

def run_tests():
    # Create Test_Results directory if it doesn't exist
    if not os.path.exists('Test_Results'):
        os.makedirs('Test_Results')
    
    # Generate timestamp for the output file
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"Test_Results/test_results_{timestamp}.txt"
    
    # Discover all test files in the 'tests' directory
    test_files = []
    for file in os.listdir('tests'):
        if file.startswith('test_') and file.endswith('.py'):
            test_files.append(file[:-3])  # Remove .py extension
    
    # Run tests from each file and collect results
    with open(output_file, 'w') as f:
        total_tests = 0
        passed_tests = 0
        failed_tests = 0
        
        for test_file in test_files:
            f.write(f"\n===== Testing {test_file} =====\n")
            module = import_module(f'tests.{test_file}')
            
            # Find all test classes in the module
            test_classes = []
            for name, obj in getmembers(module):
                if isinstance(obj, type) and issubclass(obj, unittest.TestCase):
                    test_classes.append(obj)
            
            for test_class in test_classes:
                suite = unittest.TestLoader().loadTestsFromTestCase(test_class)
                result = unittest.TextTestRunner(stream=f, verbosity=2).run(suite)
                
                total_tests += result.testsRun
                passed_tests += result.testsRun - len(result.failures) - len(result.errors)
                failed_tests += len(result.failures) + len(result.errors)
        
        # Write summary
        f.write("\n===== TEST SUMMARY =====\n")
        f.write(f"Total Tests Run: {total_tests}\n")
        f.write(f"Passed: {passed_tests}\n")
        f.write(f"Failed: {failed_tests}\n")
        f.write(f"Success Rate: {passed_tests/max(total_tests, 1)*100:.2f}%\n")
    
    print(f"Test results written to {output_file}")

if __name__ == '__main__':
    run_tests()