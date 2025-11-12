This folder holds project-local Civil 3D templates seeded into each new project:
- Discipline templates (.dwt): Survey + Site Design (CTB/STB; Imperial/Metric as needed)
- Plan Production & Sheet templates (.dwt)
- Optional reference files: USACE .arg profile, PC3/CTB/STB, icons

Authoritative Civil 3D Standards Root is defined in:
PROJECTS/_Organizer/standards.json  (key: StandardsRoot)

In Civil 3D:
1) Import the USACE .arg profile and point paths to the StandardsRoot
2) Set Gravity/Pressure pipe catalogs to:
   ${StandardsRoot}/Pipes Catalog
   ${StandardsRoot}/Pressure Pipes Catalog
