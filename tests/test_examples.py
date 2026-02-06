"""
Example tests demonstrating how to use the pytest fixtures.

These examples show patterns for writing additional tests for the SPC generator.
"""
import pytest
import numpy as np
import pandas as pd
import os
from spc_generator.generator import (
    _calculate_within_std_dev,
    sanitize_filename,
    get_unique_filepath,
    process_single_file
)


# =============================================================================
# EXAMPLE 1: Using data fixtures for unit tests
# =============================================================================

def test_statistical_calculation_with_fixture(normal_distribution_data):
    """
    Example: Test a statistical calculation using pre-generated data.

    This demonstrates using the normal_distribution_data fixture for
    testing functions that perform statistical calculations.
    """
    # The fixture provides 100 normally distributed values
    mean = np.mean(normal_distribution_data)
    std = np.std(normal_distribution_data, ddof=1)

    # Test that our data is reasonably normal
    assert 9.5 < mean < 10.5, "Mean should be around 10.0"
    assert 0.3 < std < 0.7, "Std dev should be around 0.5"


def test_subgroup_calculation(subgroup_test_data):
    """
    Example: Test subgroup-based calculations.

    This demonstrates using the subgroup_test_data fixture which
    contains 10 subgroups of 5 samples each.
    """
    # Calculate within-subgroup standard deviation
    result = _calculate_within_std_dev(subgroup_test_data, 5)

    # Verify the calculation worked
    assert result is not None, "Should successfully calculate with valid subgroups"
    assert result > 0, "Standard deviation must be positive"
    assert isinstance(result, float), "Result should be a float"


# =============================================================================
# EXAMPLE 2: Testing edge cases with edge_case_data fixture
# =============================================================================

def test_empty_data_handling(edge_case_data):
    """
    Example: Test how functions handle edge cases.

    The edge_case_data fixture provides several problematic datasets
    that functions should handle gracefully.
    """
    # Test with empty array
    empty_result = _calculate_within_std_dev(edge_case_data['empty'], 5)
    assert empty_result is None, "Should return None for empty data"

    # Test with constant values (zero variance)
    constant_std = np.std(edge_case_data['constant'], ddof=1)
    assert constant_std == 0.0, "Constant data should have zero std deviation"

    # Test with single value
    single_result = _calculate_within_std_dev(edge_case_data['single'], 5)
    assert single_result is None, "Should return None for insufficient data"


@pytest.mark.parametrize("case_name", ['empty', 'single'])
def test_edge_cases_parametrized(edge_case_data, case_name):
    """
    Example: Parametrized test using edge case data.

    This runs the same test with multiple different edge cases.
    """
    data = edge_case_data[case_name]
    result = _calculate_within_std_dev(data, 5)

    # These cases should return None (insufficient/invalid data)
    assert result is None, f"Should return None for {case_name} data"


# =============================================================================
# EXAMPLE 3: Using the Excel file creation fixture
# =============================================================================

def test_custom_excel_file_creation(create_spc_excel_file):
    """
    Example: Create custom Excel files for testing.

    The create_spc_excel_file fixture is a factory that can create
    Excel files with custom feature data.
    """
    # Define custom features for testing
    custom_features = [
        {
            'name': 'TestFeature_A',
            'nominal': 100.0,
            'usl': 105.0,
            'lsl': 95.0,
            'subgroup': 5,
            'samples': [100.1, 100.2, 99.9, 100.0, 100.3, 99.8, 100.1, 99.9, 100.2, 100.0]
        },
        {
            'name': 'TestFeature_B',
            'nominal': 50.0,
            'usl': 51.0,
            'lsl': 49.0,
            'subgroup': 3,
            'samples': [50.1, 49.9, 50.0, 50.2, 49.8, 50.0]
        }
    ]

    # Create the Excel file
    filepath = create_spc_excel_file(
        filename="SPC-DATA_custom_test.xlsx",
        features=custom_features
    )

    # Verify the file exists and can be read
    assert os.path.exists(filepath)

    # Read back and verify structure
    df = pd.read_excel(filepath, header=0, skiprows=6)
    assert 'TestFeature_A' in df.columns
    assert 'TestFeature_B' in df.columns


def test_default_excel_file(create_spc_excel_file):
    """
    Example: Use default Excel file creation.

    Without arguments, the fixture creates a file with default test features.
    """
    filepath = create_spc_excel_file()

    assert os.path.exists(filepath)
    assert filepath.endswith('.xlsx')

    # Verify it has the standard structure
    df = pd.read_excel(filepath, header=0, skiprows=6)
    assert len(df.columns) >= 2, "Should have at least 2 feature columns"


# =============================================================================
# EXAMPLE 4: Testing file system operations with temp_dir
# =============================================================================

def test_file_operations_with_temp_dir(temp_dir):
    """
    Example: Test file operations using temporary directory.

    The temp_dir fixture provides a clean temporary directory that
    is automatically cleaned up after the test.
    """
    # Create a test file
    test_file = os.path.join(temp_dir, "test_output.txt")

    with open(test_file, 'w') as f:
        f.write("test content")

    # Verify file operations
    assert os.path.exists(test_file)

    with open(test_file, 'r') as f:
        content = f.read()
        assert content == "test content"

    # No cleanup needed - fixture handles it automatically


def test_unique_filepath_generation(temp_dir):
    """
    Example: Test unique filename generation.

    Demonstrates testing file naming logic with temporary files.
    """
    # Create initial file
    base_path = os.path.join(temp_dir, "output.xlsx")
    open(base_path, 'w').close()

    # Get unique path
    unique_path = get_unique_filepath(base_path)

    # Should have added _1 suffix
    expected = os.path.join(temp_dir, "output_1.xlsx")
    assert unique_path == expected


# =============================================================================
# EXAMPLE 5: Testing with multiple fixtures
# =============================================================================

def test_with_multiple_fixtures(
    sample_metadata,
    sample_feature_data,
    temp_dir
):
    """
    Example: Combine multiple fixtures in one test.

    You can use as many fixtures as needed by adding them as parameters.
    """
    # Verify metadata fixture
    assert "Part Number" in sample_metadata
    assert sample_metadata["Part Number"] == "PN-12345"

    # Verify feature data fixture
    assert len(sample_feature_data) == 3
    first_feature = sample_feature_data[0]
    assert first_feature['name'] == 'Diameter'

    # Use temp_dir for file operations
    test_file = os.path.join(temp_dir, "combined_test.txt")
    with open(test_file, 'w') as f:
        f.write(f"Testing {first_feature['name']}")

    assert os.path.exists(test_file)


# =============================================================================
# EXAMPLE 6: Integration test with markers
# =============================================================================

@pytest.mark.integration
@pytest.mark.slow
def test_full_processing_pipeline(create_spc_excel_file, temp_dir):
    """
    Example: Integration test for full processing pipeline.

    This is marked as 'integration' and 'slow' so it can be skipped
    during fast test runs: pytest -m "not slow"
    """
    # Create input file
    input_file = create_spc_excel_file(filename="SPC-DATA_integration.xlsx")

    # Note: Uncomment to actually test process_single_file
    # process_single_file(input_file)

    # # Verify output was created
    # output_file = input_file.replace('SPC-DATA_', 'SPC-RESULTS_')
    # assert os.path.exists(output_file), "Output file should be created"

    # # Verify output structure
    # wb = load_workbook(output_file)
    # assert 'SPC_Charts' in wb.sheetnames

    pass  # Remove this when uncommenting above


# =============================================================================
# EXAMPLE 7: Fixture with custom scope
# =============================================================================

@pytest.fixture(scope="module")
def expensive_calculation():
    """
    Example: Module-scoped fixture for expensive operations.

    This fixture runs once per module instead of once per test,
    useful for expensive setup operations.
    """
    # Simulate expensive calculation
    result = sum(range(1000000))
    return result


def test_using_expensive_fixture_1(expensive_calculation):
    """First test using the module-scoped fixture."""
    assert expensive_calculation > 0


def test_using_expensive_fixture_2(expensive_calculation):
    """Second test reuses the same fixture instance."""
    assert expensive_calculation == sum(range(1000000))


# =============================================================================
# EXAMPLE 8: Custom fixture composition
# =============================================================================

@pytest.fixture
def feature_with_outliers(sample_feature_data):
    """
    Example: Create a derived fixture from an existing one.

    This takes sample_feature_data and adds outliers to test
    robustness of calculations.
    """
    feature = sample_feature_data[0].copy()
    feature['data'] = feature['data'].copy()

    # Add outliers
    feature['data'] = np.append(feature['data'], [999.0, -999.0])

    return feature


def test_outlier_handling(feature_with_outliers):
    """Example: Test using the derived fixture."""
    data = feature_with_outliers['data']

    # Verify outliers are present
    assert np.max(data) > 100
    assert np.min(data) < -100

    # Test that calculations still work
    mean = np.mean(data)
    assert not np.isnan(mean)


# =============================================================================
# EXAMPLE 9: Conditional test skipping
# =============================================================================

@pytest.mark.skipif(
    os.name != 'nt',
    reason="Windows-specific test"
)
def test_windows_specific_behavior(temp_dir):
    """Example: Test that only runs on Windows."""
    # Windows-specific file path testing
    path = os.path.join(temp_dir, "test.txt")
    assert '\\' in path or '/' in path


# =============================================================================
# EXAMPLE 10: Expected failures for known issues
# =============================================================================

@pytest.mark.xfail(reason="Known limitation: doesn't handle subgroup size > 15")
def test_large_subgroup_size():
    """
    Example: Mark test as expected to fail.

    Useful for documenting known limitations.
    """
    data = np.random.normal(10, 1, 100)
    result = _calculate_within_std_dev(data, 20)
    assert result is not None  # This will fail, but that's expected
