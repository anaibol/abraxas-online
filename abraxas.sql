-- phpMyAdmin SQL Dump
-- version 2.10.3
-- http://www.phpmyadmin.net
-- 
-- Servidor: localhost
-- Tiempo de generación: 08-06-2012 a las 20:40:50
-- Versión del servidor: 5.0.51
-- Versión de PHP: 5.2.6

SET SQL_MODE="NO_AUTO_VALUE_ON_ZERO";

-- 
-- Base de datos: `abraxas`
-- 

-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `blacksmitharmors`
-- 

CREATE TABLE `blacksmitharmors` (
  `id` int(11) NOT NULL,
  `ObjId` tinyint(3) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 
-- Volcar la base de datos para la tabla `blacksmitharmors`
-- 


-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `blacksmithhelms`
-- 

CREATE TABLE `blacksmithhelms` (
  `id` int(11) NOT NULL,
  `ObjId` tinyint(3) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 
-- Volcar la base de datos para la tabla `blacksmithhelms`
-- 


-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `blacksmithshields`
-- 

CREATE TABLE `blacksmithshields` (
  `id` int(11) NOT NULL,
  `ObjId` tinyint(3) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 
-- Volcar la base de datos para la tabla `blacksmithshields`
-- 


-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `blacksmithweapons`
-- 

CREATE TABLE `blacksmithweapons` (
  `id` int(11) NOT NULL,
  `ObjId` tinyint(3) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 
-- Volcar la base de datos para la tabla `blacksmithweapons`
-- 


-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `carpentry`
-- 

CREATE TABLE `carpentry` (
  `id` tinyint(2) NOT NULL,
  `objId` tinyint(3) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 
-- Volcar la base de datos para la tabla `carpentry`
-- 


-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `clases`
-- 

CREATE TABLE `clases` (
  `id` tinyint(2) NOT NULL,
  `nombre` varchar(15) default NULL,
  PRIMARY KEY  (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- 
-- Volcar la base de datos para la tabla `clases`
-- 

INSERT INTO `clases` VALUES (1, 'Mago');
INSERT INTO `clases` VALUES (2, 'Cl&#233;rigo');
INSERT INTO `clases` VALUES (3, 'Guerrero');
INSERT INTO `clases` VALUES (4, 'Asesino');
INSERT INTO `clases` VALUES (5, 'Ladr&#243;n');
INSERT INTO `clases` VALUES (6, 'Bardo');
INSERT INTO `clases` VALUES (7, 'Druida');
INSERT INTO `clases` VALUES (8, 'Bandido');
INSERT INTO `clases` VALUES (9, 'Palad&#237;n');
INSERT INTO `clases` VALUES (10, 'Cazador');
INSERT INTO `clases` VALUES (11, 'Pirata');

-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `guildas`
-- 

CREATE TABLE `guildas` (
  `id` tinyint(3) unsigned NOT NULL,
  `nombre` varchar(15) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- 
-- Volcar la base de datos para la tabla `guildas`
-- 


-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `history`
-- 

CREATE TABLE `history` (
  `id` smallint(4) unsigned NOT NULL auto_increment,
  `date` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `people` smallint(3) unsigned NOT NULL,
  `record` smallint(3) unsigned NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=7295 ;

-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `items`
-- 

CREATE TABLE `items` (
  `id` smallint(6) NOT NULL COMMENT 'Identifier of the object',
  `name` varchar(255) NOT NULL COMMENT 'Name',
  `price` int(11) NOT NULL default '0' COMMENT 'Price object is bought for',
  `objtype` tinyint(3) unsigned NOT NULL COMMENT 'Object type (see Server.Declares for OBJTYPE_ consts)',
  `weapontype` tinyint(3) unsigned NOT NULL COMMENT 'Weapon type (Only valid if obj=weapon - see Server.Declares)',
  `weaponrange` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Range of the weapon''s attack (if ranged)',
  `classreq` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Only allow certain classes to use this item (0 for no req)',
  `grhindex` int(11) NOT NULL COMMENT 'Index of the object graphic (by Grh value)',
  `usegrh` int(11) NOT NULL default '0' COMMENT 'Grh for the weapon''s attack',
  `usesfx` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Sound played when the object is used (0 for none)',
  `projectilerotatespeed` tinyint(3) unsigned NOT NULL COMMENT 'If a projectile, how fast it rotates (0 for no rotate)',
  `stacking` smallint(6) NOT NULL default '-1' COMMENT 'Amount the item can be stacked ( < 1 for server limit)',
  `sprite_body` smallint(6) NOT NULL default '-1' COMMENT 'Paperdolling body changed to upon usage (-1 for no change)',
  `sprite_weapon` smallint(6) NOT NULL default '-1' COMMENT 'Paperdolling weapon changed to upon usage (-1 for no change)',
  `sprite_hair` smallint(6) NOT NULL default '-1' COMMENT 'Paperdolling hair changed to upon usage (-1 for no change)',
  `sprite_head` smallint(6) NOT NULL default '-1' COMMENT 'Paperdolling head changed to upon usage (-1 for no change)',
  `sprite_wings` smallint(6) NOT NULL default '-1' COMMENT 'Paperdolling wings changed to upon usage (-1 for no change)',
  `replenish_hp` int(11) NOT NULL default '0' COMMENT 'Amount of HP replenished upon usage',
  `replenish_mp` int(11) NOT NULL default '0' COMMENT 'Amount of MP replenished upon usage',
  `replenish_sp` int(11) NOT NULL default '0' COMMENT 'Amount of SP replenished upon usage',
  `replenish_hp_percent` int(11) NOT NULL default '0' COMMENT 'Percent of HP replenished upon usage',
  `replenish_mp_percent` int(11) NOT NULL default '0' COMMENT 'Percent of MP replenished upon usage',
  `replenish_sp_percent` int(11) NOT NULL default '0' COMMENT 'Percent of SP replenished upon usage',
  `stat_str` int(11) NOT NULL default '0' COMMENT 'Strength raised upon usage',
  `stat_agi` int(11) NOT NULL default '0' COMMENT 'Agility raised upon usage',
  `stat_mag` int(11) NOT NULL default '0' COMMENT 'Magic raised upon usage',
  `stat_def` int(11) NOT NULL default '0' COMMENT 'Defence raised upon usage',
  `stat_speed` int(11) NOT NULL default '0' COMMENT 'Walk speed raised upon usage',
  `stat_hit_min` int(11) NOT NULL default '0' COMMENT 'Minimum hit raised upon usage',
  `stat_hit_max` int(11) NOT NULL default '0' COMMENT 'Maximum hit raised upon usage',
  `stat_hp` int(11) NOT NULL default '0' COMMENT 'Health raised upon usage',
  `stat_mp` int(11) NOT NULL default '0' COMMENT 'Magic raised upon usage',
  `stat_sp` int(11) NOT NULL default '0' COMMENT 'Stamina raised upon usage',
  `req_str` int(10) unsigned NOT NULL default '0' COMMENT 'Required strength to use the item',
  `req_agi` int(10) unsigned NOT NULL default '0' COMMENT 'Required agility to use the item',
  `req_mag` int(10) unsigned NOT NULL default '0' COMMENT 'Required magic to use the item',
  `req_lvl` int(10) unsigned NOT NULL default '0' COMMENT 'Required level to use the item',
  PRIMARY KEY  (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- 
-- Volcar la base de datos para la tabla `items`
-- 


-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `nombres`
-- 

CREATE TABLE `nombres` (
  `id` int(6) NOT NULL auto_increment,
  `nombre` varchar(15) NOT NULL,
  PRIMARY KEY  (`id`)
) ENGINE=MyISAM  DEFAULT CHARSET=utf8 AUTO_INCREMENT=8342 ;

-- 
-- Volcar la base de datos para la tabla `nombres`
-- 

INSERT INTO `nombres` VALUES (1, 'Aaa');

-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `npcs`
-- 

CREATE TABLE `npcs` (
  `Id` smallint(3) NOT NULL default '0' COMMENT 'Identifier of the NPC',
  `Comercia` tinyint(1) NOT NULL default '0',
  `Name` varchar(255) NOT NULL COMMENT 'Name',
  `Type` tinyint(1) NOT NULL default '0',
  `Desc` varchar(255) NOT NULL COMMENT 'Description',
  `Chat` tinyint(3) unsigned NOT NULL default '0' COMMENT 'Index of the NPC chat from the NPC Chat file',
  `Attackable` tinyint(3) unsigned NOT NULL default '0' COMMENT 'If the NPC is attackable (1 = yes, 0 = no)',
  `Hostile` tinyint(3) unsigned NOT NULL default '0' COMMENT 'If the NPC is hostile (1 = yes, 0 = no)',
  `Quest` smallint(6) NOT NULL default '0' COMMENT 'ID of the quest the NPC gives',
  `Drops` mediumtext NOT NULL COMMENT 'List of NPC drops',
  `Exp` int(11) NOT NULL default '0' COMMENT 'Experience given upon killing the NPC',
  `Objs` mediumtext NOT NULL COMMENT 'Objects sold as a shopkeeper/vendor',
  `Head` smallint(6) NOT NULL default '1' COMMENT 'Paperdolling head ID',
  `Body` smallint(6) NOT NULL default '1' COMMENT 'Paperdolling body ID',
  `Def` int(11) NOT NULL default '0' COMMENT 'Defence',
  `MinHit` int(11) NOT NULL default '1' COMMENT 'Minimum hit',
  `MaxHit` int(11) NOT NULL default '1' COMMENT 'Maximum hit',
  `MinHp` int(11) NOT NULL default '10' COMMENT 'Health points',
  `MinSta` int(11) NOT NULL default '10' COMMENT 'Stamina points',
  `Domable` tinyint(2) NOT NULL default '0',
  `Backup` tinyint(1) NOT NULL default '0',
  `Respawn` tinyint(1) NOT NULL default '0',
  `Heading` tinyint(1) NOT NULL default '0',
  `Movement` tinyint(1) NOT NULL default '0',
  `TipoItems` tinyint(1) NOT NULL default '0',
  `MinELV` tinyint(2) NOT NULL default '0',
  `MaxELV` tinyint(2) NOT NULL default '0',
  `AguaValida` tinyint(1) NOT NULL default '0',
  `DefM` tinyint(2) NOT NULL default '0',
  `Evasion` tinyint(2) NOT NULL default '0',
  `Ataque` tinyint(2) NOT NULL default '0',
  `Snd1` tinyint(3) NOT NULL default '0',
  `Snd2` tinyint(3) NOT NULL default '0',
  `Snd3` tinyint(3) NOT NULL default '0',
  `Veneno` tinyint(1) NOT NULL default '0',
  `LanzaSpells` tinyint(1) NOT NULL default '0',
  `Spells` tinytext NOT NULL,
  PRIMARY KEY  (`Id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- 
-- Volcar la base de datos para la tabla `npcs`
-- 


-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `people`
-- 

CREATE TABLE `people` (
  `Id` mediumint(6) unsigned NOT NULL auto_increment,
  `Name` varchar(15) NOT NULL,
  `Pass` varchar(30) NOT NULL,
  `Email` varchar(30) NOT NULL,
  `Act_Code` varchar(7) NOT NULL,
  `Raza` tinyint(1) unsigned NOT NULL,
  `Clase` tinyint(2) unsigned NOT NULL,
  `Genero` tinyint(1) unsigned NOT NULL,
  `Hogar` tinyint(2) unsigned NOT NULL,
  `Desc` varchar(50) NOT NULL,
  `Head` smallint(3) unsigned NOT NULL,
  `Fuerza` tinyint(2) unsigned NOT NULL,
  `Agilidad` tinyint(2) unsigned NOT NULL,
  `Inteligencia` tinyint(2) unsigned NOT NULL,
  `Carisma` tinyint(2) unsigned NOT NULL,
  `Constitucion` tinyint(2) unsigned NOT NULL,
  `ELV` tinyint(3) unsigned NOT NULL,
  `Exp` int(9) unsigned NOT NULL,
  `Skills` text NOT NULL,
  `FreeSkills` smallint(4) unsigned NOT NULL,
  `Spells` text NOT NULL,
  `Map` smallint(3) unsigned NOT NULL,
  `X` tinyint(3) unsigned NOT NULL,
  `Y` tinyint(3) unsigned NOT NULL,
  `MinHP` smallint(4) unsigned NOT NULL,
  `MaxHP` smallint(4) unsigned NOT NULL,
  `MinMan` smallint(5) unsigned NOT NULL,
  `MaxMan` smallint(5) unsigned NOT NULL,
  `MinSta` smallint(4) unsigned NOT NULL,
  `MaxSta` smallint(4) unsigned NOT NULL,
  `MinHit` smallint(3) unsigned NOT NULL,
  `MaxHit` smallint(3) unsigned NOT NULL,
  `MinSed` tinyint(3) unsigned NOT NULL,
  `MinHam` tinyint(3) unsigned NOT NULL,
  `Matados` smallint(5) unsigned NOT NULL,
  `NpcMatados` smallint(5) unsigned NOT NULL,
  `Muertes` smallint(5) unsigned NOT NULL,
  `Inv` text,
  `Belt` text,
  `Bank` text,
  `Gld` int(9) unsigned NOT NULL,
  `BankGld` int(9) unsigned NOT NULL,
  `HeadEqp` tinyint(4) unsigned NOT NULL,
  `BodyEqp` tinyint(4) unsigned NOT NULL,
  `LeftHandEqp` text NOT NULL,
  `RightHandEqp` text NOT NULL,
  `AmmoAmount` text NOT NULL,
  `BeltEqp` text NOT NULL,
  `RingEqp` text NOT NULL,
  `Ship` text NOT NULL,
  `Compas` text NOT NULL,
  `Mascos` text NOT NULL,
  `Plataformas` text NOT NULL,
  `Guild_Id` tinyint(3) unsigned default NULL,
  `Envenenado` tinyint(1) NOT NULL,
  `Silencio` tinyint(1) NOT NULL,
  `Pena_Carcel` tinyint(1) NOT NULL,
  `Ban` tinyint(1) NOT NULL,
  `Logged` tinyint(1) NOT NULL,
  `Date_Created` datetime default NULL,
  `Last_Ip` varchar(15) NOT NULL,
  `Last_Ip2` varchar(15) NOT NULL,
  `Last_Ip3` varchar(15) NOT NULL,
  `Last_Date` datetime default NULL,
  `Last_Date2` datetime default NULL,
  `Last_Date3` datetime default NULL,
  `Last_Update` timestamp NOT NULL default '0000-00-00 00:00:00' on update CURRENT_TIMESTAMP,
  `UpTime` int(9) unsigned NOT NULL,
  `Priv` tinyint(1) NOT NULL default '1',
  PRIMARY KEY  (`Id`),
  KEY `Name` (`Name`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=2414 ;

-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `platforms`
-- 

CREATE TABLE `platforms` (
  `id` tinyint(2) NOT NULL,
  `map` tinyint(3) NOT NULL,
  `x` tinyint(2) NOT NULL,
  `y` tinyint(2) NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 
-- Volcar la base de datos para la tabla `platforms`
-- 


-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `razas`
-- 

CREATE TABLE `razas` (
  `Id` varchar(1) NOT NULL,
  `Nombre` varchar(15) NOT NULL,
  PRIMARY KEY  (`Id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

-- 
-- Volcar la base de datos para la tabla `razas`
-- 

INSERT INTO `razas` VALUES ('1', 'Humano');
INSERT INTO `razas` VALUES ('2', 'Elfo');
INSERT INTO `razas` VALUES ('3', 'Elfo oscuro');
INSERT INTO `razas` VALUES ('4', 'Gnomo');
INSERT INTO `razas` VALUES ('5', 'Enano');

-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `spells`
-- 

CREATE TABLE `spells` (
  `id` tinyint(3) NOT NULL,
  `name` varchar(50) collate utf8_unicode_ci NOT NULL,
  `desc` varchar(50) collate utf8_unicode_ci NOT NULL,
  `magicWords` varchar(50) collate utf8_unicode_ci NOT NULL,
  `hechizeroMsg` varchar(50) collate utf8_unicode_ci NOT NULL,
  `propioMsg` varchar(50) collate utf8_unicode_ci NOT NULL,
  `targetMsg` varchar(50) collate utf8_unicode_ci NOT NULL,
  `type` tinyint(3) NOT NULL default '0',
  `snd` tinyint(3) NOT NULL default '0',
  `fxGrh` smallint(5) NOT NULL default '0',
  `fxLoops` tinyint(1) NOT NULL default '0',
  `minSkill` tinyint(3) NOT NULL default '0',
  `manaRequired` tinyint(4) NOT NULL default '0',
  `staRequired` tinyint(4) NOT NULL default '0',
  `targetType` tinyint(1) NOT NULL default '0',
  `subeHP` tinyint(1) NOT NULL default '0',
  `minHP` tinyint(4) NOT NULL default '0',
  `maxHP` tinyint(4) NOT NULL default '0',
  `subeMana` tinyint(1) NOT NULL default '0',
  `minMana` tinyint(4) NOT NULL default '0',
  `maxMana` tinyint(4) NOT NULL default '0',
  `subeSta` tinyint(1) NOT NULL default '0',
  `minSta` tinyint(4) NOT NULL default '0',
  `maxSta` tinyint(4) NOT NULL default '0',
  `subeHam` tinyint(1) NOT NULL default '0',
  `minHam` tinyint(3) NOT NULL default '0',
  `maxHam` tinyint(3) NOT NULL default '0',
  `subeSed` tinyint(1) NOT NULL default '0',
  `minSed` tinyint(3) NOT NULL default '0',
  `maxSed` tinyint(3) NOT NULL default '0',
  `subeAg` tinyint(1) NOT NULL default '0',
  `minAg` tinyint(2) NOT NULL default '0',
  `maxAg` int(2) NOT NULL default '0',
  `subeFu` tinyint(1) NOT NULL default '0',
  `minFu` tinyint(2) NOT NULL default '0',
  `maxFu` tinyint(2) NOT NULL default '0',
  `subeCa` tinyint(1) NOT NULL default '0',
  `minCa` tinyint(2) NOT NULL default '0',
  `maxCa` tinyint(2) NOT NULL default '0',
  `invi` tinyint(1) NOT NULL default '0',
  `paraliza` tinyint(1) NOT NULL default '0',
  `inmo` tinyint(1) NOT NULL default '0',
  `remueveInmo` tinyint(1) NOT NULL default '0',
  `remueveEstupidez` tinyint(1) NOT NULL default '0',
  `remueveInviParcial` tinyint(1) NOT NULL default '0',
  `curaVeneno` tinyint(1) NOT NULL default '0',
  `envenena` tinyint(1) NOT NULL default '0',
  `revive` tinyint(1) NOT NULL default '0',
  `enceguece` tinyint(1) NOT NULL default '0',
  `estupidez` tinyint(1) NOT NULL default '0',
  `invoca` tinyint(1) NOT NULL default '0',
  `numNpc` tinyint(1) NOT NULL default '0',
  `cantidadNpc` tinyint(1) NOT NULL default '0',
  `mimetiza` tinyint(1) NOT NULL default '0',
  `materializa` tinyint(1) NOT NULL default '0',
  `itemIndex` tinyint(1) NOT NULL default '0',
  `staffAffected` tinyint(1) NOT NULL default '0',
  `needStaff` tinyint(1) NOT NULL default '0',
  `resistencia` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;

-- 
-- Volcar la base de datos para la tabla `spells`
-- 


-- --------------------------------------------------------

-- 
-- Estructura de tabla para la tabla `stats`
-- 

CREATE TABLE `stats` (
  `Online_Players` tinyint(3) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- 
-- Volcar la base de datos para la tabla `stats`
-- 

INSERT INTO `stats` VALUES (1);
