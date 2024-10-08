import pytest
from UpdateLinks import pull_relative_info, create_new_link, check_link

def test_pull_relative_info():
    assert pull_relative_info("http://example.com/documents/manuals/guide.pdf") == "guide"
    assert pull_relative_info("https://docs.anl.gov/main/groups/intranet/@shared/@lms/documents/procedure/lms-proc-281.pdf") == "lms-proc-281"

def test_create_new_link():
    assert create_new_link("https://my.anl.gov/esb/view/","lms-proc-281") == "https://my.anl.gov/esb/view/lms-proc-281"

def test_check_link():
    assert check_link('https://www.w3schools.com/python/gloss_python_loop_dictionary_items.asp') == True
    assert check_link('https://www.w3schools.com/python/gloss_python_loop_dictionary_items.asp') == False