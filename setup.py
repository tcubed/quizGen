from setuptools import setup, find_packages

setup(
    name="quizgen",
    version="0.1.0",
    packages=find_packages(),
    install_requires=["numpy","pandas","python-docx"],
    #include_package_data=True,  # Ensure package data files are included
    #ackage_data={"mypackage": ["examples/*.py"]},
)
  