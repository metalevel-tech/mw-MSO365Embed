# MediaWiki Extension MSO365Embed

**The extension is tested with MW 1.38.**

MediaWiki media tag extension for Microsoft Office files, such as "docx", "docm", "xlsx", "xlsm", "pptx", "pptm",  "ppsx" and "ppsm" files by using:

* `https://view.officeapps.live.com/op/embed.aspx?src=...` or
* `https://view.officeapps.live.com/op/view.aspx?src=...`

This extension is based on [Extension:PDFEmbed](https://github.com/WolfgangFahl/PDFEmbed) and is distributed under the same license.

Examples of usage:

```xml
<mso365 width="500px" height="300px">Example.pptx</mso365>
<mso365 style="your css style">File:Example.docx</mso365>
<mso365 action='view'>File:Example.xlsx</mso365>
```

* You can omit the `File:` part. It should handle also URLs as `https://example.com/your.docx` instead of `File:Example.docx`, etc.

## Installation

Clone the extension:

```bash
cd $IP/extensions
sudo git clone https://github.com/metalevel-tech/mw-MSO365Embed.git MSO365Embed # HTTPS
sudo git clone git@github.com:metalevel-tech/mw-MSO365Embed.git MSO365Embed     # SSH
```

To install this extension, add the following to the end of the `LocalSettings.php` file:

```php
wfLoadExtension('MSO365Embed');
```

## Configuration

If the default configuration needs to be altered add these settings to the `LocalSettings.php` file below `wfLoadExtension('MSO365Embed')`:

```php
$wgMSO365Embed['height'] = '696px'; // CSS Width of the wrapper div
$wgMSO365Embed['width'] = '100%';   // CSS Height of the wrapper div
$wgMSO365Embed['style'] = 'border-radius: 0; border: 1px solid #323639; margin: 8px auto 18px;'; // CSS Style ...
$wgMSO365Embed['action'] = 'embed'; // Actions: embed | view
$wgMSO365Embed['iframe'] = false;   // 'true' use Html:iframe, 'false' (default) use Html:object
$wgGroupPermissions['*']['embed_MSO365'] = true;
```

* For the default values see [extension.json](extension.json).

## See also

* [Extension:MSO365Handler](https://github.com/metalevel-tech/mw-MSO365Handler)

## References

* <https://github.com/WolfgangFahl/PDFEmbed>
* <https://doc.wikimedia.org/mediawiki-core/master/php>
* <https://www.mediawiki.org/wiki/Manual:UserFactory.php>
* <https://www.mediawiki.org/wiki/Manual:Tag_extensions>
* <https://www.mediawiki.org/wiki/Manual:Hooks/ParserFirstCallInit>
* <https://doc.wikimedia.org/mediawiki-core/master/php/classHtml.html#a92f023b28be16bb69004084d66a8ac38>
* <https://stackoverflow.com/a/60809767/6543935>
