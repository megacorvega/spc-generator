# SPC Generator - Testing Guide

This document describes the pytest fixtures and testing infrastructure for the SPC Generator tool.

## Setup

Install test dependencies:

```bash
pip install -r requirements.txt
```

## Running Tests

### Run all tests
```bash
pytest
```

### Run with verbose output
```bash
pytest -v
```

### Run specific test file
```bash
pytest test_spc_generator.py
```

### Run specific test
```bash
pytest test_spc_generator.py::test_calculate_within_std_dev_normal_case
```

### Run with coverage report
```bash
pytest --cov=spc_generator --cov-report=html
```

### Run only fast tests (skip slow tests)
```bash
pytest -m "not slow"
```

## Available Fixtures

### Data Fixtures

#### `sample_metadata`
Provides sample metadata dictionary for SPC reports.

**Returns:** Dict with Part Number, Batch Number, Date of Inspection, Inspector

**Usage:**
```python
def test_metadata(sample_metadata):
    assert sample_metadata["Part Number"] == "PN-12345"
```

#### `sample_feature_data`
Provides a list of 3 feature data dictionaries with normally distributed data.

**Returns:** List of dicts with keys: name, nominal, usl, lsl, subgroup, data

**Usage:**
```python
def test_features(sample_feature_data):
    for feature in sample_feature_data:
        assert len(feature['data']) > 0
```

#### `minimal_feature_data`
Feature with only 1 data point (edge case testing).

**Usage:**
```python
def test_minimal_data(minimal_feature_data):
    assert len(minimal_feature_data['data']) == 1
```

#### `insufficient_subgroup_data`
Feature with < 2 subgroups (for testing within-variation edge cases).

**Usage:**
```python
def test_insufficient_subgroups(insufficient_subgroup_data):
    result = _calculate_within_std_dev(
        insufficient_subgroup_data['data'],
        insufficient_subgroup_data['subgroup']
    )
    assert result is None
```

#### `normal_distribution_data`
100 data points from normal distribution (mean=10, std=0.5).

**Usage:**
```python
def test_statistics(normal_distribution_data):
    mean = np.mean(normal_distribution_data)
    assert 9.8 < mean < 10.2
```

#### `subgroup_test_data`
50 data points arranged in 10 subgroups of 5, specifically designed for subgroup testing.

**Usage:**
```python
def test_subgroup_calculation(subgroup_test_data):
    result = _calculate_within_std_dev(subgroup_test_data, 5)
    assert result is not None
```

#### `edge_case_data`
Dictionary of edge case datasets: empty, single, constant, two_points, with_nan, outliers.

**Usage:**
```python
def test_empty_data(edge_case_data):
    empty_array = edge_case_data['empty']
    assert len(empty_array) == 0
```

### File System Fixtures

#### `temp_dir`
Creates a temporary directory for test files (auto-cleanup after test).

**Returns:** Path to temporary directory

**Usage:**
```python
def test_file_creation(temp_dir):
    test_file = os.path.join(temp_dir, "test.txt")
    with open(test_file, 'w') as f:
        f.write("test")
    assert os.path.exists(test_file)
```

#### `create_spc_excel_file`
Factory fixture to create properly formatted SPC Excel input files.

**Returns:** Function that creates Excel files

**Usage:**
```python
def test_excel_processing(create_spc_excel_file):
    # Create with default features
    filepath = create_spc_excel_file()

    # Or create with custom features
    custom_features = [
        {
            'name': 'CustomFeature',
            'nominal': 50.0,
            'usl': 52.0,
            'lsl': 48.0,
            'subgroup': 5,
            'samples': [50.1, 49.9, 50.2, 49.8, 50.0]
        }
    ]
    filepath = create_spc_excel_file("custom.xlsx", features=custom_features)

    assert os.path.exists(filepath)
```

#### `duplicate_file_setup`
Creates files for testing unique filepath generation.

**Returns:** Dict with directory, base_name, expected_next

**Usage:**
```python
def test_unique_naming(duplicate_file_setup):
    next_path = get_unique_filepath(
        os.path.join(duplicate_file_setup['directory'], 'test_file.xlsx')
    )
    assert next_path == duplicate_file_setup['expected_next']
```

### Configuration Fixtures

#### `c4_constant_tests`
Test cases for c4 constant validation.

**Returns:** Dict mapping subgroup sizes to expected c4 values

**Usage:**
```python
def test_c4_lookup(c4_constant_tests):
    for size, expected in c4_constant_tests.items():
        actual = C4_CONSTANTS.get(size)
        assert actual == expected
```

#### `filename_sanitization_cases`
Test cases for filename sanitization (tuples of input/expected output).

**Usage:**
```python
def test_sanitization(filename_sanitization_cases):
    for original, expected in filename_sanitization_cases:
        result = sanitize_filename(original)
        assert '/' not in result and '\\' not in result
```

### Helper Fixtures (from conftest.py)

#### `mock_input`
Factory to mock input() calls.

**Usage:**
```python
def test_user_input(mock_input):
    mock_input(['yes', 'no'])
    # Your code that calls input()
```

#### `no_sleep`
Mocks time.sleep() to speed up tests with delays.

**Usage:**
```python
def test_with_delay(no_sleep):
    # time.sleep() calls will be instant
    wait_for_file_access('some_file.xlsx')
```

#### `reset_matplotlib`
Automatically resets matplotlib state between tests (autouse=True).

#### `test_environment`
Provides Python version, platform, and working directory info (session-scoped).

## Writing New Tests

### Example: Testing a calculation function

```python
def test_my_calculation(normal_distribution_data):
    """Test description."""
    result = my_function(normal_distribution_data)
    assert result > 0
    assert isinstance(result, float)
```

### Example: Testing Excel file processing

```python
@pytest.mark.integration
def test_process_file(create_spc_excel_file):
    """Test full file processing."""
    filepath = create_spc_excel_file()

    # Process the file
    process_single_file(filepath)

    # Check output was created
    output_path = filepath.replace('SPC-DATA_', 'SPC-RESULTS_')
    assert os.path.exists(output_path)
```

### Example: Testing edge cases

```python
def test_edge_cases(edge_case_data):
    """Test function handles edge cases."""
    # Empty data
    result = my_function(edge_case_data['empty'])
    assert result is None

    # Constant data
    result = my_function(edge_case_data['constant'])
    assert result == 0  # Zero variance
```

## Test Markers

Use markers to categorize tests:

```python
@pytest.mark.unit
def test_unit_function():
    pass

@pytest.mark.integration
def test_integration():
    pass

@pytest.mark.slow
def test_slow_operation():
    pass
```

Run specific categories:
```bash
pytest -m unit          # Only unit tests
pytest -m integration   # Only integration tests
pytest -m "not slow"    # Skip slow tests
```

## Coverage Reports

Generate HTML coverage report:
```bash
pytest --cov=spc_generator --cov-report=html
```

View report:
```bash
# Opens in browser (Windows)
start htmlcov/index.html
```

## Troubleshooting

### Tests fail with "ModuleNotFoundError"
Ensure you've installed requirements:
```bash
pip install -r requirements.txt
```

### Tests fail with matplotlib errors
The conftest.py file should handle this automatically by setting the 'Agg' backend.
If issues persist, verify matplotlib is properly installed.

### Temporary files not cleaning up
The `temp_dir` fixture includes automatic cleanup. If files persist, check for:
- Tests that don't finish (exceptions)
- File locks (close all Excel files)

## Best Practices

1. **Use existing fixtures** when possible instead of creating test data in each test
2. **Name tests clearly** with descriptive names that explain what's being tested
3. **Test one thing** per test function
4. **Use markers** to organize tests by type/speed
5. **Clean up resources** (fixtures handle this automatically)
6. **Mock external dependencies** (file I/O, user input, etc.)
7. **Test edge cases** using the provided edge_case_data fixture

## Next Steps

To expand test coverage:

1. Add integration tests for `process_single_file()`
2. Add tests for plotting functions (`create_summary_image`, etc.)
3. Add parameterized tests for various data distributions
4. Add tests for file locking behavior (`wait_for_file_access`)
5. Add tests for Excel output formatting and styling
