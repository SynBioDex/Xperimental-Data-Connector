from setuptools import find_packages, setup

setup(name='xperimental-data-conv',
      version='1.0.1b',
      url='https://github.com/SynBioDex/Experimental-Data-Convertor',
      license='BSD 3-clause',
      maintainer='Gonzalo Vidal',
      maintainer_email='Gonzalo.vidalpena@colorado.edu',
      include_package_data=True,
      description='Convert Excel resources into SBOL and Flapjack, uploads them to SynBioHub and Flapjack and connects them',
      packages=find_packages(include=['xperimental_data_conv']),
      long_description=open('README.md').read(),
      install_requires=['excel2flapjack==1.0.8',
                        'excel2sbol==1.0.29'],
      zip_safe=False)
