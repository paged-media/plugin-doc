/*
 * This file is part of paged (https://paged.media).
 *
 * paged is free software: you may redistribute it and/or modify it under the
 * terms of the GNU Affero General Public License, version 3, as published by
 * the Free Software Foundation, OR under the Paged Media Enterprise License
 * (PMEL), a commercial license available from And The Next GmbH. Full
 * copyright and license information is available in LICENSE.md, distributed
 * with this source code.
 *
 * paged is distributed in the hope that it will be useful, but WITHOUT ANY
 * WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the licenses for details.
 *
 *  @copyright  Copyright (c) And The Next GmbH
 *  @license    AGPL-3.0-only OR Paged Media Enterprise License (PMEL)
 */

//! Word's symbol characters (`<w:sym w:font=".." w:char=".."/>`) as Unicode.
//!
//! Word writes a character from a symbol font (Symbol, Wingdings) as the
//! font's own code in the private-use area: `F0xx` in a symbol font is code
//! `xx` of that font (a plain `xx`, below `0x100`, means the same). The
//! character the engine can carry is that glyph's Unicode equivalent:
//!
//! - **Symbol**: every glyph of Word's own `symbol.ttf` whose Adobe glyph
//!   name has a Unicode value (the Adobe Glyph List; Delta, Omega and mu as
//!   Adobe's `symbol.txt` maps them, the bracket and arrow pieces to their
//!   Miscellaneous Technical homes). Only `radicalex` (the radical's bar
//!   extender) has none.
//! - **Wingdings**: the glyphs whose Unicode equivalent is unambiguous,
//!   checked against the glyph names in Word's `Wingdings.ttf`, and where
//!   Word's own PDF names a glyph (`fixtures/symbols.word.json`), that name
//!   (the bold check mark is U+2713, not U+2714). The rest
//!   (envelopes, leaves, most arrows, the Windows logo, …) have no agreed
//!   equivalent and are not guessed.
//!
//! Anything else in the private-use area (another symbol font, an unmapped
//! code) has no character; a `w:char` outside it, in any font, IS the
//! character. The equivalent is drawn in the run's own font (with the
//! engine's fallback), not in the symbol font.

/// The character a `<w:sym>` stands for: `font` is `@w:font`, `code` the
/// `@w:char` hex. `None` when there is no Unicode equivalent (or the code
/// is not hex).
pub fn symbol_char(font: Option<&str>, code: &str) -> Option<char> {
    let value = u32::from_str_radix(code.trim(), 16).ok()?;
    let font_code = match value {
        0xF020..=0xF0FF => Some((value - 0xF000) as u8),
        0x20..=0xFF => Some(value as u8),
        _ => None,
    };
    let table = match font.map(|f| f.trim().to_ascii_lowercase()).as_deref() {
        Some("symbol") => Some(SYMBOL),
        Some("wingdings") => Some(WINGDINGS),
        _ => None,
    };
    if let Some(table) = table {
        let code = font_code?;
        return table
            .binary_search_by_key(&code, |(c, _)| *c)
            .ok()
            .map(|i| table[i].1);
    }
    // Not a symbol font this table knows: a private-use code names a glyph
    // of that font only; anything else is the Unicode character itself.
    if (0xE000..=0xF8FF).contains(&value) || value < 0x20 {
        return None;
    }
    char::from_u32(value)
}

/// Word's `symbol.ttf` (code → Unicode), sorted by code.
const SYMBOL: &[(u8, char)] = &[
    (0x20, '\u{0020}'), // space
    (0x21, '\u{0021}'), // exclam
    (0x22, '\u{2200}'), // universal
    (0x23, '\u{0023}'), // numbersign
    (0x24, '\u{2203}'), // existential
    (0x25, '\u{0025}'), // percent
    (0x26, '\u{0026}'), // ampersand
    (0x27, '\u{220B}'), // suchthat
    (0x28, '\u{0028}'), // parenleft
    (0x29, '\u{0029}'), // parenright
    (0x2A, '\u{2217}'), // asteriskmath
    (0x2B, '\u{002B}'), // plus
    (0x2C, '\u{002C}'), // comma
    (0x2D, '\u{2212}'), // minus
    (0x2E, '\u{002E}'), // period
    (0x2F, '\u{002F}'), // slash
    (0x30, '\u{0030}'), // zero
    (0x31, '\u{0031}'), // one
    (0x32, '\u{0032}'), // two
    (0x33, '\u{0033}'), // three
    (0x34, '\u{0034}'), // four
    (0x35, '\u{0035}'), // five
    (0x36, '\u{0036}'), // six
    (0x37, '\u{0037}'), // seven
    (0x38, '\u{0038}'), // eight
    (0x39, '\u{0039}'), // nine
    (0x3A, '\u{003A}'), // colon
    (0x3B, '\u{003B}'), // semicolon
    (0x3C, '\u{003C}'), // less
    (0x3D, '\u{003D}'), // equal
    (0x3E, '\u{003E}'), // greater
    (0x3F, '\u{003F}'), // question
    (0x40, '\u{2245}'), // congruent
    (0x41, '\u{0391}'), // Alpha
    (0x42, '\u{0392}'), // Beta
    (0x43, '\u{03A7}'), // Chi
    (0x44, '\u{0394}'), // Delta
    (0x45, '\u{0395}'), // Epsilon
    (0x46, '\u{03A6}'), // Phi
    (0x47, '\u{0393}'), // Gamma
    (0x48, '\u{0397}'), // Eta
    (0x49, '\u{0399}'), // Iota
    (0x4A, '\u{03D1}'), // theta1
    (0x4B, '\u{039A}'), // Kappa
    (0x4C, '\u{039B}'), // Lambda
    (0x4D, '\u{039C}'), // Mu
    (0x4E, '\u{039D}'), // Nu
    (0x4F, '\u{039F}'), // Omicron
    (0x50, '\u{03A0}'), // Pi
    (0x51, '\u{0398}'), // Theta
    (0x52, '\u{03A1}'), // Rho
    (0x53, '\u{03A3}'), // Sigma
    (0x54, '\u{03A4}'), // Tau
    (0x55, '\u{03A5}'), // Upsilon
    (0x56, '\u{03C2}'), // sigma1
    (0x57, '\u{03A9}'), // Omega
    (0x58, '\u{039E}'), // Xi
    (0x59, '\u{03A8}'), // Psi
    (0x5A, '\u{0396}'), // Zeta
    (0x5B, '\u{005B}'), // bracketleft
    (0x5C, '\u{2234}'), // therefore
    (0x5D, '\u{005D}'), // bracketright
    (0x5E, '\u{22A5}'), // perpendicular
    (0x5F, '\u{005F}'), // underscore
    (0x61, '\u{03B1}'), // alpha
    (0x62, '\u{03B2}'), // beta
    (0x63, '\u{03C7}'), // chi
    (0x64, '\u{03B4}'), // delta
    (0x65, '\u{03B5}'), // epsilon
    (0x66, '\u{03C6}'), // phi
    (0x67, '\u{03B3}'), // gamma
    (0x68, '\u{03B7}'), // eta
    (0x69, '\u{03B9}'), // iota
    (0x6A, '\u{03D5}'), // phi1
    (0x6B, '\u{03BA}'), // kappa
    (0x6C, '\u{03BB}'), // lambda
    (0x6D, '\u{03BC}'), // mu
    (0x6E, '\u{03BD}'), // nu
    (0x6F, '\u{03BF}'), // omicron
    (0x70, '\u{03C0}'), // pi
    (0x71, '\u{03B8}'), // theta
    (0x72, '\u{03C1}'), // rho
    (0x73, '\u{03C3}'), // sigma
    (0x74, '\u{03C4}'), // tau
    (0x75, '\u{03C5}'), // upsilon
    (0x76, '\u{03D6}'), // omega1
    (0x77, '\u{03C9}'), // omega
    (0x78, '\u{03BE}'), // xi
    (0x79, '\u{03C8}'), // psi
    (0x7A, '\u{03B6}'), // zeta
    (0x7B, '\u{007B}'), // braceleft
    (0x7C, '\u{007C}'), // bar
    (0x7D, '\u{007D}'), // braceright
    (0x7E, '\u{223C}'), // similar
    (0xA1, '\u{03D2}'), // Upsilon1
    (0xA2, '\u{2032}'), // minute
    (0xA3, '\u{2264}'), // lessequal
    (0xA4, '\u{2044}'), // fraction
    (0xA5, '\u{221E}'), // infinity
    (0xA6, '\u{0192}'), // florin
    (0xA7, '\u{2663}'), // club
    (0xA8, '\u{2666}'), // diamond
    (0xA9, '\u{2665}'), // heart
    (0xAA, '\u{2660}'), // spade
    (0xAB, '\u{2194}'), // arrowboth
    (0xAC, '\u{2190}'), // arrowleft
    (0xAD, '\u{2191}'), // arrowup
    (0xAE, '\u{2192}'), // arrowright
    (0xAF, '\u{2193}'), // arrowdown
    (0xB0, '\u{00B0}'), // degree
    (0xB1, '\u{00B1}'), // plusminus
    (0xB2, '\u{2033}'), // second
    (0xB3, '\u{2265}'), // greaterequal
    (0xB4, '\u{00D7}'), // multiply
    (0xB5, '\u{221D}'), // proportional
    (0xB6, '\u{2202}'), // partialdiff
    (0xB7, '\u{2022}'), // bullet
    (0xB8, '\u{00F7}'), // divide
    (0xB9, '\u{2260}'), // notequal
    (0xBA, '\u{2261}'), // equivalence
    (0xBB, '\u{2248}'), // approxequal
    (0xBC, '\u{2026}'), // ellipsis
    (0xBD, '\u{23D0}'), // arrowvertex
    (0xBE, '\u{23AF}'), // arrowhorizex
    (0xBF, '\u{21B5}'), // carriagereturn
    (0xC0, '\u{2135}'), // aleph
    (0xC1, '\u{2111}'), // Ifraktur
    (0xC2, '\u{211C}'), // Rfraktur
    (0xC3, '\u{2118}'), // weierstrass
    (0xC4, '\u{2297}'), // circlemultiply
    (0xC5, '\u{2295}'), // circleplus
    (0xC6, '\u{2205}'), // emptyset
    (0xC7, '\u{2229}'), // intersection
    (0xC8, '\u{222A}'), // union
    (0xC9, '\u{2283}'), // propersuperset
    (0xCA, '\u{2287}'), // reflexsuperset
    (0xCB, '\u{2284}'), // notsubset
    (0xCC, '\u{2282}'), // propersubset
    (0xCD, '\u{2286}'), // reflexsubset
    (0xCE, '\u{2208}'), // element
    (0xCF, '\u{2209}'), // notelement
    (0xD0, '\u{2220}'), // angle
    (0xD1, '\u{2207}'), // gradient
    (0xD2, '\u{00AE}'), // registerserif
    (0xD3, '\u{00A9}'), // copyrightserif
    (0xD4, '\u{2122}'), // trademarkserif
    (0xD5, '\u{220F}'), // product
    (0xD6, '\u{221A}'), // radical
    (0xD7, '\u{22C5}'), // dotmath
    (0xD8, '\u{00AC}'), // logicalnot
    (0xD9, '\u{2227}'), // logicaland
    (0xDA, '\u{2228}'), // logicalor
    (0xDB, '\u{21D4}'), // arrowdblboth
    (0xDC, '\u{21D0}'), // arrowdblleft
    (0xDD, '\u{21D1}'), // arrowdblup
    (0xDE, '\u{21D2}'), // arrowdblright
    (0xDF, '\u{21D3}'), // arrowdbldown
    (0xE0, '\u{25CA}'), // lozenge
    (0xE1, '\u{2329}'), // angleleft
    (0xE2, '\u{00AE}'), // registersans
    (0xE3, '\u{00A9}'), // copyrightsans
    (0xE4, '\u{2122}'), // trademarksans
    (0xE5, '\u{2211}'), // summation
    (0xE6, '\u{239B}'), // parenlefttp
    (0xE7, '\u{239C}'), // parenleftex
    (0xE8, '\u{239D}'), // parenleftbt
    (0xE9, '\u{23A1}'), // bracketlefttp
    (0xEA, '\u{23A2}'), // bracketleftex
    (0xEB, '\u{23A3}'), // bracketleftbt
    (0xEC, '\u{23A7}'), // bracelefttp
    (0xED, '\u{23A8}'), // braceleftmid
    (0xEE, '\u{23A9}'), // braceleftbt
    (0xEF, '\u{23AA}'), // braceex
    (0xF1, '\u{232A}'), // angleright
    (0xF2, '\u{222B}'), // integral
    (0xF3, '\u{2320}'), // integraltp
    (0xF4, '\u{23AE}'), // integralex
    (0xF5, '\u{2321}'), // integralbt
    (0xF6, '\u{239E}'), // parenrighttp
    (0xF7, '\u{239F}'), // parenrightex
    (0xF8, '\u{23A0}'), // parenrightbt
    (0xF9, '\u{23A4}'), // bracketrighttp
    (0xFA, '\u{23A5}'), // bracketrightex
    (0xFB, '\u{23A6}'), // bracketrightbt
    (0xFC, '\u{23AB}'), // bracerighttp
    (0xFD, '\u{23AC}'), // bracerightmid
    (0xFE, '\u{23AD}'), // bracerightbt
];

/// Word's `Wingdings.ttf` (code → Unicode) where the equivalent is
/// unambiguous, sorted by code. The comment is the font's glyph name.
const WINGDINGS: &[(u8, char)] = &[
    (0x22, '\u{2702}'),  // scissors
    (0x23, '\u{2701}'),  // scissorscutting
    (0x25, '\u{1F514}'), // bell
    (0x26, '\u{1F4D6}'), // book
    (0x28, '\u{260E}'),  // telephonesolid
    (0x29, '\u{2706}'),  // telhandsetcirc
    (0x30, '\u{1F4C1}'), // folder
    (0x31, '\u{1F4C2}'), // folderopen
    (0x36, '\u{231B}'),  // hourglass
    (0x37, '\u{2328}'),  // keyboard
    (0x38, '\u{1F5B1}'), // mouse2button
    (0x3E, '\u{2707}'),  // tapereel
    (0x3F, '\u{270D}'),  // handwrite
    (0x41, '\u{270C}'),  // handv
    (0x42, '\u{1F44C}'), // handok
    (0x43, '\u{1F44D}'), // thumbup
    (0x44, '\u{1F44E}'), // thumbdown
    (0x45, '\u{261C}'),  // handptleft
    (0x46, '\u{261E}'),  // handptright
    (0x47, '\u{261D}'),  // handptup
    (0x48, '\u{261F}'),  // handptdwn
    (0x49, '\u{270B}'),  // handhalt
    (0x4A, '\u{263A}'),  // smileface
    (0x4B, '\u{1F610}'), // neutralface
    (0x4C, '\u{2639}'),  // frownface
    (0x4D, '\u{1F4A3}'), // bomb
    (0x4E, '\u{2620}'),  // skullcrossbones
    (0x51, '\u{2708}'),  // airplane
    (0x52, '\u{263C}'),  // sunshine
    (0x53, '\u{1F4A7}'), // droplet
    (0x54, '\u{2744}'),  // snowflake
    (0x56, '\u{271E}'),  // crossshadow
    (0x58, '\u{2720}'),  // crossmaltese
    (0x59, '\u{2721}'),  // starofdavid
    (0x5A, '\u{262A}'),  // crescentstar
    (0x5B, '\u{262F}'),  // yinyang
    (0x5C, '\u{0950}'),  // om
    (0x5D, '\u{2638}'),  // wheel
    (0x5E, '\u{2648}'),  // aries
    (0x5F, '\u{2649}'),  // taurus
    (0x60, '\u{264A}'),  // gemini
    (0x61, '\u{264B}'),  // cancer
    (0x62, '\u{264C}'),  // leo
    (0x63, '\u{264D}'),  // virgo
    (0x64, '\u{264E}'),  // libra
    (0x65, '\u{264F}'),  // scorpio
    (0x66, '\u{2650}'),  // saggitarius
    (0x67, '\u{2651}'),  // capricorn
    (0x68, '\u{2652}'),  // aquarius
    (0x69, '\u{2653}'),  // pisces
    (0x6C, '\u{25CF}'),  // circle6
    (0x6D, '\u{274D}'),  // circleshadowdwn
    (0x6E, '\u{25A0}'),  // square6
    (0x6F, '\u{25A1}'),  // box3
    (0x71, '\u{2751}'),  // boxshadowdwn
    (0x72, '\u{2752}'),  // boxshadowup
    (0x73, '\u{2B27}'),  // lozenge4
    (0x74, '\u{29EB}'),  // lozenge6
    (0x75, '\u{25C6}'),  // rhombus6
    (0x76, '\u{2756}'),  // xrhombus
    (0x77, '\u{2B25}'),  // rhombus4
    (0x78, '\u{2327}'),  // clear
    (0x7A, '\u{2318}'),  // command
    (0x7B, '\u{2740}'),  // rosette
    (0x7C, '\u{273F}'),  // rosettesolid
    (0x7D, '\u{275D}'),  // quotedbllftbld
    (0x7E, '\u{275E}'),  // quotedblrtbld
    (0x80, '\u{24EA}'),  // zerosans
    (0x81, '\u{2460}'),  // onesans
    (0x82, '\u{2461}'),  // twosans
    (0x83, '\u{2462}'),  // threesans
    (0x84, '\u{2463}'),  // foursans
    (0x85, '\u{2464}'),  // fivesans
    (0x86, '\u{2465}'),  // sixsans
    (0x87, '\u{2466}'),  // sevensans
    (0x88, '\u{2467}'),  // eightsans
    (0x89, '\u{2468}'),  // ninesans
    (0x8A, '\u{2469}'),  // tensans
    (0x8B, '\u{24FF}'),  // zerosansinv
    (0x8C, '\u{2776}'),  // onesansinv
    (0x8D, '\u{2777}'),  // twosansinv
    (0x8E, '\u{2778}'),  // threesansinv
    (0x8F, '\u{2779}'),  // foursansinv
    (0x90, '\u{277A}'),  // fivesansinv
    (0x91, '\u{277B}'),  // sixsansinv
    (0x92, '\u{277C}'),  // sevensansinv
    (0x93, '\u{277D}'),  // eightsansinv
    (0x94, '\u{277E}'),  // ninesansinv
    (0x95, '\u{277F}'),  // tensansinv
    (0x9E, '\u{00B7}'),  // circle2
    (0x9F, '\u{2022}'),  // circle4
    (0xA1, '\u{25CB}'),  // ring2
    (0xA4, '\u{25C9}'),  // ringbutton2
    (0xA5, '\u{25CE}'),  // target
    (0xA7, '\u{25AA}'),  // square4
    (0xA8, '\u{25FB}'),  // box2
    (0xAA, '\u{2726}'),  // crosstar2
    (0xAB, '\u{2605}'),  // pentastar2
    (0xAC, '\u{2736}'),  // hexstar2
    (0xAD, '\u{2734}'),  // octastar2
    (0xAE, '\u{2739}'),  // dodecastar3
    (0xAF, '\u{2735}'),  // octastar4
    (0xB0, '\u{2316}'),  // registersquare
    (0xB4, '\u{2370}'),  // query
    (0xB5, '\u{272A}'),  // circlestar
    (0xB6, '\u{2730}'),  // starshadow
    (0xB7, '\u{1F550}'), // oneoclock
    (0xB8, '\u{1F551}'), // twooclock
    (0xB9, '\u{1F552}'), // threeoclock
    (0xBA, '\u{1F553}'), // fouroclock
    (0xBB, '\u{1F554}'), // fiveoclock
    (0xBC, '\u{1F555}'), // sixoclock
    (0xBD, '\u{1F556}'), // sevenoclock
    (0xBE, '\u{1F557}'), // eightoclock
    (0xBF, '\u{1F558}'), // nineoclock
    (0xC0, '\u{1F559}'), // tenoclock
    (0xC1, '\u{1F55A}'), // elevenoclock
    (0xC2, '\u{1F55B}'), // twelveoclock
    (0xD5, '\u{232B}'),  // deleteleft
    (0xD6, '\u{2326}'),  // deleteright
    (0xD8, '\u{27A2}'),  // head2right
    (0xEF, '\u{21E6}'),  // bleft
    (0xF0, '\u{21E8}'),  // bright
    (0xF1, '\u{21E7}'),  // bup
    (0xF2, '\u{21E9}'),  // bdown
    (0xF3, '\u{2B04}'),  // bleftright
    (0xF4, '\u{21F3}'),  // bupdown
    (0xF5, '\u{2B01}'),  // bnw
    (0xF6, '\u{2B00}'),  // bne
    (0xF7, '\u{2B03}'),  // bsw
    (0xF8, '\u{2B02}'),  // bse
    (0xFB, '\u{2718}'),  // xmarkbld
    (0xFC, '\u{2713}'),  // checkbld (Word's own PDF names it U+2713)
    (0xFD, '\u{2612}'),  // boxxmarkbld
    (0xFE, '\u{2611}'),  // boxcheckbld
];

#[cfg(test)]
mod tests {
    #![allow(non_snake_case)] // `…__feat__<id>` test names link the cockpit feature
    use super::*;

    #[test]
    fn tables_are_sorted_for_the_binary_search__feat__plugin_doc_read_path() {
        for t in [SYMBOL, WINGDINGS] {
            assert!(t.windows(2).all(|w| w[0].0 < w[1].0));
        }
    }

    #[test]
    fn symbol_and_wingdings_codes_map_to_their_characters__feat__plugin_doc_read_path() {
        assert_eq!(symbol_char(Some("Symbol"), "F0B7"), Some('\u{2022}'));
        assert_eq!(symbol_char(Some("Symbol"), "B7"), Some('\u{2022}'));
        assert_eq!(symbol_char(Some("Symbol"), "F061"), Some('\u{03B1}'));
        assert_eq!(symbol_char(Some("Symbol"), "F044"), Some('\u{0394}'));
        assert_eq!(symbol_char(Some("Wingdings"), "F0FC"), Some('\u{2713}'));
        assert_eq!(symbol_char(Some("wingdings"), "F0A7"), Some('\u{25AA}'));
        assert_eq!(symbol_char(Some("Wingdings"), "F0D8"), Some('\u{27A2}'));
    }

    #[test]
    fn codes_without_an_equivalent_have_no_character__feat__plugin_doc_read_path() {
        // The radical extender; the Windows logo; an unknown symbol font.
        assert_eq!(symbol_char(Some("Symbol"), "F060"), None);
        assert_eq!(symbol_char(Some("Wingdings"), "F0FF"), None);
        assert_eq!(symbol_char(Some("Webdings"), "F021"), None);
        assert_eq!(symbol_char(Some("Symbol"), "zz"), None);
        // Outside the private-use area the code IS the character.
        assert_eq!(symbol_char(Some("Arial"), "2014"), Some('\u{2014}'));
        assert_eq!(symbol_char(None, "00E9"), Some('\u{00E9}'));
    }
}
