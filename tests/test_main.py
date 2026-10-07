from my_project.main import greet


def test_greet_default():
    assert greet() == "Hello, world!"


def test_greet_name():
    assert greet("Ann") == "Hello, Ann!"
