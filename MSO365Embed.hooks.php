<?php

/**
 * MSO365Embed
 * MSO365Embed  Hooks
 *
 * @author		Spas Z. Spasov (source of the idea: https://github.com/WolfgangFahl/PDFEmbed)
 * @license		LGPLv3 http://opensource.org/licenses/lgpl-3.0.html
 * @package		MSO365Embed
 * @link		https://github.com/metalevel-tech/mw-MSO365Embed
 *
 */

class MSO365Embed
{

    /**
     *
     * Sets up this extensions parser functions.
     *
     * @access public
     * @param
     *            object Parser object passed as a reference.
     * @return boolean true
     */
    static public function onParserFirstCallInit(Parser &$parser)
    {
        $parser->setHook('MSO365', 'MSO365Embed::generateTag');
        return true;
    }

    /**
     * disable the cache, (probably we don't need it, because the file is cached by office.live for few hours)
     *
     * @param Parser $parser
     */
    static public function disableCache(Parser &$parser)
    {
        // see https://gerrit.wikimedia.org/r/plugins/gitiles/mediawiki/extensions/MagicNoCache/+/refs/heads/master/src/MagicNoCacheHooks.php
        global $wgOut;
        $parser->getOutput()->updateCacheExpiry(0);
        $wgOut->disableClientCache();
    }

    /**
     * remove the File: prefix depending on the language or in english default form
     *
     * @param
     *            filename - the filename for which to fix the prefix
     *
     * @return    string - the filename without the File: / Media: or i18n File/Media prefix
     */
    static public function removeFilePrefix($filename)
    {
        $mwServices = MediaWiki\MediaWikiServices::getInstance();
        if (method_exists($mwServices, "getContentLanguage")) {
            $contentLang = $mwServices->getContentLanguage();
            # there are four possible prefixes
            $ns_media_wiki_lang = $contentLang->getFormattedNsText(NS_MEDIA);
            $ns_file_wiki_lang  = $contentLang->getFormattedNsText(NS_FILE);
            if (method_exists($mwServices, "getLanguageFactory")) {
                $langFactory = $mwServices->getLanguageFactory();
                $lang = $langFactory->getLanguage('en');
                $ns_media_lang_en = $lang->getFormattedNsText(NS_MEDIA);
                $ns_file_lang_en  = $lang->getFormattedNsText(NS_FILE);
                $filename = preg_replace("/^($ns_media_wiki_lang|$ns_file_wiki_lang|$ns_media_lang_en|$ns_file_lang_en):/", '', $filename);
            } else {
                $filename = preg_replace("/^($ns_media_wiki_lang|$ns_file_wiki_lang):/", '', $filename);
            }
        }
        return $filename;
    }

    /**
     * Generates the MSO365/mso365 object tag.
     *
     * @access public
     *
     * @param
     *            string Namespace prefixed article of the MSO365 file to display.
     * @param
     *            array Arguments on the tag.
     * @param
     *            object Parser object.
     * @param
     *            object PPFrame object.
     *
     * @return string HTML
     */

    static public function generateTag($input, array $args, Parser $parser, PPFrame $frame)
    {
        global $wgMSO365Embed, $wgRequest;
        // disable the cache
        MSO365Embed::disableCache($parser);

        // grab the uri by parsing to html
        $html = $parser->recursiveTagParse($input);

        // check the action which triggered us
        $requestAction = $wgRequest->getVal('action');

        // depending on the action get the responsible user
        if ($requestAction == 'edit' || $requestAction == 'submit') {
            $user = RequestContext::getMain()->getUser();
        } else {
            // https://www.mediawiki.org/wiki/Manual:UserFactory.php
            $revUserName = $parser->getRevisionUser();
            $userFactory = MediaWiki\MediaWikiServices::getInstance()->getUserFactory();
            $user = $userFactory->newFromName($revUserName);
        }

        if ($user === false) {
            return self::error('embed_mso365_invalid_user');
        }

        if (!MediaWiki\MediaWikiServices::getInstance()->getPermissionManager()->userHasRight($user, 'embed_MSO365')) {
            return self::error('embed_mso365_no_permission');
        }

        // we don't want the html but just the href of the link
        // so we might reverse some of the parsing again by examining the html
        // whether it contains an anchor <a href= ...
        if (strpos($html, '<a') !== false) {
            $a = new SimpleXMLElement($html);
            // is there a href element?
            if (isset($a['href'])) {
                // that's what we want ...
                $html = $a['href'];
            }
        }

        // Handle the arguments of the current <mso365> tag
        (isset($args['style']))  ? $style = $args['style']   : $style = $wgMSO365Embed['style'];
        (isset($args['width']))  ? $width = $args['width']   : $width = $wgMSO365Embed['width'];
        (isset($args['height'])) ? $height = $args['height'] : $height = $wgMSO365Embed['height'];
        (isset($args['action'])) ? $action = $args['action'] : $action = $wgMSO365Embed['action'];
        (isset($args['iframe'])) ? $iframe = $args['iframe'] : $iframe = $wgMSO365Embed['iframe'];


        // if there are no slashes in the name we assume this might be a pointer to a file
        if (preg_match('~^([^\/]+\.(docx|docm|xlsx|xlsm|pptm|pptx|ppsx|ppsm))(#[0-9]+)?$~', $html, $re)) {
            // re contains the groups
            $filename = $re[1];

            $filename = self::removeFilePrefix($filename);
            $MSO365File =  MediaWiki\MediaWikiServices::getInstance()->getRepoGroup()->findFile($filename);

            if ($MSO365File !== false) {
                $url = $MSO365File->getFullUrl();
                return self::embed($url, $width, $height, $style, $action, $iframe);
            } else {
                return self::error('embed_mso365_invalid_file', $filename);
            }
        } else {
            // parse the given url
            $domain = parse_url($html);

            // check that the parsing worked and retrieve a valid host
            // no relative urls are allowed ...
            if ($domain === false || (!isset($domain['host']))) {
                if (!isset($domain['host'])) {
                    return self::error("embed_mso365_invalid_relative_domain: ", $html);
                }
                return self::error("embed_mso365_invalid_url", $html);
            }

            // check that url is valid
            if (filter_var($html, FILTER_VALIDATE_URL, FILTER_FLAG_PATH_REQUIRED)) {
                return self::embed($html, $width, $height, $style, $action, $iframe);
            } else {
                return self::error('embed_mso365_invalid_url', $html);
            }
        }
    }

    /**
     * Returns an HTML node for the given file as string.
     *
     * @access private
     *
     * @param
     *            URL url to embed.
     * @param
     *            integer width of the iframe.
     * @param
     *            integer height of the iframe.
     * @param
     *            integer style of the div container.
     *
     * @return string HTML code for iframe.
     */
    static private function embed($url, $width, $height, $style, $action, $iframe)
    {
        $divStyle = 'position: relative; overflow: hidden; width: ' . $width . '; height: ' . $height . '; ' . $style;
        $iframeStyle = 'position: absolute; left: -.2%; top: -.2%;';

        // compose url
        $FileUrl = htmlentities($url);
        $ViewOfficeLiveUrl = 'https://view.officeapps.live.com/op/' . $action . '.aspx?src=';
        $MSO365SafeUrl = $ViewOfficeLiveUrl . $FileUrl;

        // return a proper HTML element

        if ($iframe) {
            $output = Html::rawElement('iframe', [
                'class' => 'mso365-iframe',
                'width' => '100.4%',
                'height' => '100.4%',
                'src' => $MSO365SafeUrl,
                'frameborder' => '0',
                'style' => $iframeStyle
            ]);
        } else {
            $output = Html::rawElement('object', [
                'class' => 'mso365-iframe',
                'width' => '100.4%',
                'height' => '100.4%',
                'data' => $MSO365SafeUrl,
                'type' => 'application/xml',
                'frameborder' => '0',
                'style' => $iframeStyle
            ]);
        }


        return Html::rawElement('div', [
            'class' => 'mso365-div',
            'style' => $divStyle
        ], $output);
    }

    /**
     * Returns a standard error message.
     *
     * @access private
     * @param
     *            string Error message key to display.
     * @param
     *            params any parameters for the error message
     * @return string HTML error message.
     */
    static private function error($messageKey, ...$params)
    {
        return Xml::span(wfMessage($messageKey, $params)->plain(), 'error');
    }
}
