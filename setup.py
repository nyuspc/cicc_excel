import setuptools

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setuptools.setup(
    name="cicc_excel",
    version="0.0.2",
    author="Pengcheng Song",
    author_email="smth_spc@hotmail.com",
    description="Library of export pandas into excel for CICCers",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/nyuspc/cicc_excel",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python",
        "Intended Audience :: Financial and Insurance Industry",
        "Topic :: Scientific/Engineering :: Mathematics",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)