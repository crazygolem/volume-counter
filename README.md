Google Apps Script and Spreadsheet bindings to count the number of volumes in a
book inventory.

The syntax is human-readable and accomodates (lists of) numbered or named
volumes, ranges of numbered volumes, specific counts, and groups of volumes:

| List of volumes                                    | Count |
| -------------------------------------------------- | ----- |
| 1-10, 15-18                                        | 14    |
| 00, 1, 1.5, 2-4                                    | 6     |
| 1-6, Collector [1-3], 2 * Art Book                 | 11    |
| Collector [Sleeve [1-3], Sleeve [4-6], 2 * Extras] | 8     |
