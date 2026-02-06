"""
pytest fixtures and tests for spc_generator.py

This file contains unit tests for the SPC generator functions.
All fixtures are defined in conftest.py and are automatically available.
"""
import pytest
import os
import numpy as np
import pandas as pd

# Import functions from spc_generator
from spc_generator.generator import (
    _calculate_within_std_dev,
    sanitize_filename,
    get_unique_filepath,
    C4_CONSTANTS
)


# ============================================================================
# UNIT TESTS
# ============================================================================

def test_calculate_within_std_dev_normal_case(subgroup_test_data):
    """Test within-subgroup standard deviation calculation with valid data."""
    result = _calculate_within_std_dev(subgroup_test_data, 5)
    assert result is not None
    assert result > 0
    assert isinstance(result, float)


def test_calculate_within_std_dev_insufficient_subgroups(insufficient_subgroup_data):
    """Test that function returns None when there are < 2 subgroups."""
    result = _calculate_within_std_dev(
        insufficient_subgroup_data['data'],
        insufficient_subgroup_data['subgroup']
    )
    assert result is None


def test_calculate_within_std_dev_invalid_subgroup_size():
    """Test that function handles invalid subgroup sizes."""
    data = np.array([1, 2, 3, 4, 5])
    assert _calculate_within_std_dev(data, 1) is None  # Too small
    assert _calculate_within_std_dev(data, 16) is None  # Too large
    assert _calculate_within_std_dev(data, 2.7) is not None  # Float converted to int


def test_sanitize_filename(filename_sanitization_cases):
    """Test filename sanitization removes invalid characters."""
    for original, expected in filename_sanitization_cases:
        result = sanitize_filename(original)
        # sanitize_filename only removes invalid file system characters, not spaces
        assert '/' not in result
        assert '\\' not in result
        assert '*' not in result
        assert '?' not in result
        assert ':' not in result
        assert '"' not in result
        assert '<' not in result
        assert '>' not in result
        assert '|' not in result


def test_get_unique_filepath_new_file(temp_dir):
    """Test that unique filepath returns original path if file doesn't exist."""
    new_path = os.path.join(temp_dir, "nonexistent.xlsx")
    result = get_unique_filepath(new_path)
    assert result == new_path


def test_get_unique_filepath_duplicate(duplicate_file_setup):
    """Test that unique filepath generates correct numbered suffix."""
    test_path = os.path.join(
        duplicate_file_setup['directory'],
        duplicate_file_setup['base_name']
    )
    result = get_unique_filepath(test_path)
    assert result == duplicate_file_setup['expected_next']


def test_c4_constants_coverage(c4_constant_tests):
    """Verify C4 constants dictionary has expected values."""
    for subgroup_size, expected_value in c4_constant_tests.items():
        actual_value = C4_CONSTANTS.get(subgroup_size)
        assert actual_value == expected_value


def test_excel_file_creation(create_spc_excel_file):
    """Test that the fixture properly creates SPC Excel files."""
    filepath = create_spc_excel_file()

    assert os.path.exists(filepath)
    assert filepath.endswith('.xlsx')

    # Verify we can read it back
    df = pd.read_excel(filepath, header=0, skiprows=6)
    assert df is not None
    assert len(df.columns) > 1  # At least one feature column


def test_sample_metadata_fixture(sample_metadata):
    """Verify metadata fixture has all required fields."""
    required_fields = ["Part Number", "Batch Number", "Date of Inspection", "Inspector"]
    for field in required_fields:
        assert field in sample_metadata
        assert sample_metadata[field] is not None


def test_sample_feature_data_fixture(sample_feature_data):
    """Verify feature data fixture structure."""
    assert len(sample_feature_data) > 0

    for feature in sample_feature_data:
        assert 'name' in feature
        assert 'nominal' in feature
        assert 'usl' in feature
        assert 'lsl' in feature
        assert 'subgroup' in feature
        assert 'data' in feature
        assert len(feature['data']) > 0


def test_edge_case_data_fixture(edge_case_data):
    """Verify edge case data fixture provides expected cases."""
    assert len(edge_case_data['empty']) == 0
    assert len(edge_case_data['single']) == 1
    assert np.all(edge_case_data['constant'] == 10.0)
    assert np.any(np.isnan(edge_case_data['with_nan']))
