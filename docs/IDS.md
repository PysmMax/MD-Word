# IDS.md — GUID та версії пакетів (Фаза 0)

Зафіксовано один раз у Фазі 0. **Не змінювати** без явного рішення — обидва
GUID зашиті в майбутню COM-реєстрацію (Фаза 3) та інсталятор (Фаза 5).

## GUID

Згенеровано `[guid]::NewGuid()` (PowerShell) 2026-07-02, свіжі, без збігів з
попередніми спробами.

| Ім'я | GUID | Призначення |
|---|---|---|
| `CLSID_Connect` | `A18CE6CA-7818-468A-868D-69E34B4469F6` | CLSID COM-класу `MdWord.AddIn.Connect` (реєстр `HKCU\Software\Classes\CLSID\{...}`) |
| `INNO_APPID` | `B9C839C5-DB90-4C62-9F81-E9DEE82CA417` | `AppId` в Inno Setup скрипті (`MdWord.iss`) — ключ upgrade-поверх |

## Версії пакетів

Джерело: `dotnet package search <id> --take 3` (nuget.org) 2026-07-02,
підтверджено net48-сумісність за `.nuspec` (`<group targetFramework=...>`)
через `curl https://api.nuget.org/v3-flatcontainer/<id>/<ver>/<id>.nuspec`.

| Пакет | Версія | Проєкт | Причина вибору |
|---|---|---|---|
| `Microsoft.NETFramework.ReferenceAssemblies` | `1.0.*` | усі три | зафіксовано планом (wildcard дозволено явно) — targeting pack net48 відсутній локально, без цього пакета net48 не збереться |
| `Markdig` | `1.3.2` | `MdWord.Core` | latest stable на nuget.org на момент розвідки; єдиний .NETFramework-таргет пакета (net-агностичний, PCL-стиль) — сумісний з net48 |
| `DocumentFormat.OpenXml` | `3.5.1` | `MdWord.Core` | latest stable 3.x (як вимагає план); `.nuspec` містить group `.NETFramework4.6` → сумісний з net48 |
| `Jint` | `4.10.1` | `MdWord.Core` | latest stable 4.x (як вимагає план); `.nuspec` містить group `.NETFramework4.6.2` → сумісний з net48 |
| `Microsoft.Office.Interop.Word` | `15.0.4797.1004` | `MdWord.AddIn` | єдина версія, опублікована на nuget.org; репак Office 2013/2016/2019 interop-збірки без жодних залежностей (не тягне `office.dll`) → використано з `EmbedInteropTypes=true` замість плану Б (`dynamic`); **перевірено лише, що пакет резолвиться і `dotnet build` зелений** — на цьому етапі `MdWord.AddIn` не містить жодного `.cs`-файлу, тому жоден interop-тип фактично не embed-иться і не використовується; функціональна перевірка інтеропу — задача Фази 3 |
| `xunit` | `2.9.3` | `MdWord.Core.Tests` | latest stable xUnit v2 (v3 — окремий несумісний API, план не вимагає міграції) |
| `xunit.runner.visualstudio` | `3.1.5` | `MdWord.Core.Tests` | latest; опис пакета явно підтверджує сумісність із xUnit v1/v2/v3 і .NET Framework 4.7.2+ |
| `Microsoft.NET.Test.Sdk` | `18.7.0` | `MdWord.Core.Tests` | latest stable; `.nuspec` містить dependency-group `.NETFramework4.6.2` → сумісний з net48 |
| `katex` | `0.16.47` | `MdWord.Core` (embedded resource, не NuGet) | latest patch у гілці 0.16.x на момент розвідки (2026-07-03, `npm view katex versions`) — план явно фіксує 0.16.x, не 0.17.x (latest на npm); `dist/katex.min.js` витягнуто з офіційного тарболу (`npm pack katex@0.16.47`), не скопійовано вручну зі сторінки |
| `xsltml` (mmltex) | `2.1.2` | `MdWord.Core` (embedded resources, не NuGet) | MathML → LaTeX (Фаза 4, зворотний шлях). Джерело — SourceForge `xsltml` project (`sourceforge.net/projects/xsltml/files/xsltml/v.2.1.2/xsltml_2.1.2.zip`), завантажено і розпаковано напряму (не з пам'яті/переказу) 2026-07-03. **XSLT-версія перевірена емпірично**: усі 7 файлів (`mmltex.xsl` + `tokens.xsl`/`glayout.xsl`/`scripts.xsl`/`tables.xsl`/`entities.xsl`/`cmarkup.xsl`) мають `version='1.0'` на кореневому `xsl:stylesheet` — сумісні з `XslCompiledTransform` (XSLT 1.0 only), на відміну від `github.com/davidcarlisle/web-xslt`, чий MathML→LaTeX конвертер живе в іншій підпапці (`pmml2tex`, не `mmltex`) і є XSLT 2.0 — тому не використаний. **Ліцензія**: `README` дистрибутиву — "Copyright (C) 2001-2003 Vasil Yaroshevich", дозвільна MIT-подібна ліцензія (вільне використання/копіювання/модифікація/розповсюдження за умови збереження copyright notice; похідні стилі, що публічно розповсюджуються, мають бути перейменовані — не стосується внутрішнього використання без модифікації вендорених файлів). Вендорено байт-в-байт без змін (`src/MdWord.Core/Resources/mmltex/*.xsl` + `README-xsltml.txt`, усі — embedded resources); власний файл `mdword-entry.xsl` (НЕ частина дистрибутиву xsltml) через `xsl:import` перевизначає верхньошаблонний `m:math`, щоб прибрати вбудовані mmltex-делімітери (`$ ... $` / `\[...\]`) — делімітер (`$...$` чи `$$...$$`) вибирає викликач у C#, залежно від того, це інлайн `m:oMath` чи блоковий `m:oMathPara`. Багатофайловий `xsl:include`/`xsl:import` розв'язується простим шляхом (без кастомного `XmlResolver`): при першому виклику всі 8 файлів витягуються з embedded resources у тимчасову теку процесу, і `XslCompiledTransform.Load` викликається з фізичного шляху — так само, як `MML2OMML.XSL`/`OMML2MML.XSL` уже завантажуються з реальних шляхів Office, а не з ресурсів. |

## Примітка щодо Microsoft.Office.Interop.Word

План передбачав план Б (`dynamic`) як рівноцінний варіант, якщо пакет
недоступний або тягне `office.dll`. Пакет виявився доступним і без
залежностей, тож використано `PackageReference` з `EmbedInteropTypes="true"`
у каркасі `MdWord.AddIn` (без коду — це буде Фаза 3).

**Важливо:** зелений `dotnet build` на цьому етапі підтверджує лише, що
пакет резолвиться з nuget.org і не тягне зайвих залежностей — НЕ що
interop функціонально працює. `MdWord.AddIn` зараз не містить жодного
`.cs`-файлу, тому `EmbedInteropTypes` не embed-ить жодного типу (embed
відбувається лише для типів, які реально використовує код). Реальна
перевірка інтеропу (виклики Word OM, сумісність зі встановленою версією
Word) — задача Фази 3. Якщо там виявляться проблеми, рішення можна
переглянути на план Б без зміни структури проєкту.
