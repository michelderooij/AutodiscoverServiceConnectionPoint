# AutodiscoverServiceConnectionPoint

This script contains helper functions to create, remove and configure 
Autodiscover Service Connection Points for Exchange.
	
## Prerequisites

Script requires the Active Directory PowerShell Module.
	
## Usage

```
Clear-AutodiscoverServiceConnectionPoint -Name EX1
Clear-AutodiscoverServiceConnectionPoint -Name EX2
New-AutodiscoverServiceConnectionPoint -Name EX1 -ServiceBinding https://autodiscover.contoso.com/autodiscover/autodiscover.xml -Site Default-First-Site-Name
New-AutodiscoverServiceConnectionPoint -Name EX2 -ServiceBinding https://autodiscover.contoso.com/autodiscover/autodiscover.xml
Set-AutodiscoverServiceConnectionPoint -Name EX2 -Site Default-First-Site-Name
```

## Contributing

N/A

## Versioning

Initial version published on GitHub is 1.0. Changelog is contained in the script.

## Authors

* Michel de Rooij [initial work] https://github.com/michelderooij

## License

This project is licensed under the MIT License - see the LICENSE.md for details.

## Acknowledgments

N/A
 