# M0009 Real Word Tier-3 Evidence

- Platform: Microsoft Windows NT 10.0.26200.0 (x64)
- Microsoft Word version reported by automation: `16.0`
- Repository revision: `5cf39296a6943e475e3d79fa38debb58b18cb2d0` plus the working-tree M0009 changes
- Authored DOCX: `authored.docx`
- Authored DOCX SHA-256: `9245676853102901C15521105B3FEF876759A6E55DC275DDD63DB58174295667`
- Finalized DOCX: `finalized.docx`
- Finalized DOCX SHA-256: `FFFBF6B2EBE6B9B8470BF4EC2DEA15751DA3C6B7BA912BCD8C1DE0D57EDB853C`
- Command/result: `dotnet test tests/Yadg.IntegrationTests/Yadg.IntegrationTests.csproj --no-build --configuration Release --filter FullyQualifiedName~M0009IntegrationTests` passed 6/6.
- Coverage: real Word opened, refreshed, and saved a synthetic DOCX containing a numbered section, figure/table captions with `SEQ`, `REF` fields, template-owned TOC/list fields, a generated table, a figure, and ordinary content; authored input remained byte-identical, finalized output reopened successfully, and the test observed no newly owned `WINWORD.EXE` process after completion.
