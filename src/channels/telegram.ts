import fs from 'fs';
import https from 'https';
import path from 'path';

import { Api, Bot } from 'grammy';

import { ASSISTANT_NAME, TRIGGER_PATTERN } from '../config.js';
import { readEnvFile } from '../env.js';
import { resolveGroupFolderPath } from '../group-folder.js';
import { logger } from '../logger.js';
import { registerChannel, ChannelOpts } from './registry.js';
import {
  Channel,
  OnChatMetadata,
  OnInboundMessage,
  RegisteredGroup,
} from '../types.js';

/** Download a URL to a local file path using the native https module. */
function downloadFile(url: string, dest: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(dest);
    https
      .get(url, (res) => {
        if (res.statusCode !== 200) {
          file.destroy();
          fs.unlink(dest, () => {});
          reject(new Error(`HTTP ${res.statusCode} downloading file`));
          return;
        }
        res.pipe(file);
        file.on('finish', () => file.close(() => resolve()));
      })
      .on('error', (err) => {
        file.destroy();
        fs.unlink(dest, () => {});
        reject(err);
      });
  });
}

export interface TelegramChannelOpts {
  onMessage: OnInboundMessage;
  onChatMetadata: OnChatMetadata;
  registeredGroups: () => Record<string, RegisteredGroup>;
  allowedUserIds?: Set<string>;
}

/**
 * Send a message with Telegram Markdown parse mode, falling back to plain text.
 * Claude's output naturally matches Telegram's Markdown v1 format:
 *   *bold*, _italic_, `code`, ```code blocks```, [links](url)
 */
async function sendTelegramMessage(
  api: { sendMessage: Api['sendMessage'] },
  chatId: string | number,
  text: string,
  options: { message_thread_id?: number } = {},
): Promise<void> {
  try {
    await api.sendMessage(chatId, text, {
      ...options,
      parse_mode: 'Markdown',
    });
  } catch (err) {
    // Fallback: send as plain text if Markdown parsing fails
    logger.debug({ err }, 'Markdown send failed, falling back to plain text');
    await api.sendMessage(chatId, text, options);
  }
}

export class TelegramChannel implements Channel {
  name = 'telegram';

  private bot: Bot | null = null;
  private opts: TelegramChannelOpts;
  private botToken: string;

  constructor(botToken: string, opts: TelegramChannelOpts) {
    this.botToken = botToken;
    this.opts = opts;
  }

  async connect(): Promise<void> {
    this.bot = new Bot(this.botToken, {
      client: {
        baseFetchConfig: { agent: https.globalAgent, compress: true },
      },
    });

    // Command to get chat ID (useful for registration) — registered chats only
    this.bot.command('chatid', (ctx) => {
      const chatJid = `tg:${ctx.chat.id}`;
      if (!this.opts.registeredGroups()[chatJid]) return;

      const chatId = ctx.chat.id;
      const chatType = ctx.chat.type;
      const chatName =
        chatType === 'private'
          ? ctx.from?.first_name || 'Private'
          : (ctx.chat as any).title || 'Unknown';

      ctx.reply(
        `Chat ID: \`tg:${chatId}\`\nName: ${chatName}\nType: ${chatType}`,
        { parse_mode: 'Markdown' },
      );
    });

    // Command to check bot status — registered chats only
    this.bot.command('ping', (ctx) => {
      const chatJid = `tg:${ctx.chat.id}`;
      if (!this.opts.registeredGroups()[chatJid]) return;
      ctx.reply(`${ASSISTANT_NAME} is online.`);
    });

    // Telegram bot commands handled above — skip them in the general handler
    // so they don't also get stored as messages. All other /commands flow through.
    const TELEGRAM_BOT_COMMANDS = new Set(['chatid', 'ping']);

    this.bot.on('message:text', async (ctx) => {
      if (ctx.message.text.startsWith('/')) {
        const cmd = ctx.message.text.slice(1).split(/[\s@]/)[0].toLowerCase();
        if (TELEGRAM_BOT_COMMANDS.has(cmd)) return;
      }

      const chatJid = `tg:${ctx.chat.id}`;
      let content = ctx.message.text;
      const timestamp = new Date(ctx.message.date * 1000).toISOString();
      const senderName =
        ctx.from?.first_name ||
        ctx.from?.username ||
        ctx.from?.id.toString() ||
        'Unknown';
      const sender = ctx.from?.id.toString() || '';
      const msgId = ctx.message.message_id.toString();

      // Determine chat name
      const chatName =
        ctx.chat.type === 'private'
          ? senderName
          : (ctx.chat as any).title || chatJid;

      // Translate Telegram @bot_username mentions into TRIGGER_PATTERN format.
      // Telegram @mentions (e.g., @andy_ai_bot) won't match TRIGGER_PATTERN
      // (e.g., ^@Andy\b), so we prepend the trigger when the bot is @mentioned.
      const botUsername = ctx.me?.username?.toLowerCase();
      if (botUsername) {
        const entities = ctx.message.entities || [];
        const isBotMentioned = entities.some((entity) => {
          if (entity.type === 'mention') {
            const mentionText = content
              .substring(entity.offset, entity.offset + entity.length)
              .toLowerCase();
            return mentionText === `@${botUsername}`;
          }
          return false;
        });
        if (isBotMentioned && !TRIGGER_PATTERN.test(content)) {
          content = `@${ASSISTANT_NAME} ${content}`;
        }
      }

      // Store chat metadata for discovery
      const isGroup =
        ctx.chat.type === 'group' || ctx.chat.type === 'supergroup';
      this.opts.onChatMetadata(
        chatJid,
        timestamp,
        chatName,
        'telegram',
        isGroup,
      );

      // Only deliver full message for registered groups
      const group = this.opts.registeredGroups()[chatJid];
      if (!group) {
        logger.debug(
          { chatJid, chatName },
          'Message from unregistered Telegram chat',
        );
        return;
      }

      // Drop messages from unauthorized senders (if allowlist is configured)
      if (
        this.opts.allowedUserIds &&
        sender &&
        !this.opts.allowedUserIds.has(sender)
      ) {
        logger.warn(
          { chatJid, sender: senderName, userId: sender },
          'Telegram message dropped: sender not in allowlist',
        );
        return;
      }

      // Deliver message — startMessageLoop() will pick it up
      this.opts.onMessage(chatJid, {
        id: msgId,
        chat_jid: chatJid,
        sender,
        sender_name: senderName,
        content,
        timestamp,
        is_from_me: false,
      });

      logger.info(
        { chatJid, chatName, sender: senderName },
        'Telegram message stored',
      );
    });

    // Handle non-text messages with placeholders so the agent knows something was sent
    const storeNonText = (ctx: any, placeholder: string) => {
      const chatJid = `tg:${ctx.chat.id}`;
      const group = this.opts.registeredGroups()[chatJid];
      if (!group) return;

      const timestamp = new Date(ctx.message.date * 1000).toISOString();
      const senderId = ctx.from?.id?.toString() || '';
      const senderName =
        ctx.from?.first_name || ctx.from?.username || senderId || 'Unknown';
      const caption = ctx.message.caption ? ` ${ctx.message.caption}` : '';

      if (
        this.opts.allowedUserIds &&
        senderId &&
        !this.opts.allowedUserIds.has(senderId)
      ) {
        logger.warn(
          { chatJid, sender: senderName, userId: senderId },
          'Telegram message dropped: sender not in allowlist',
        );
        return;
      }

      const isGroup =
        ctx.chat.type === 'group' || ctx.chat.type === 'supergroup';
      this.opts.onChatMetadata(
        chatJid,
        timestamp,
        undefined,
        'telegram',
        isGroup,
      );
      this.opts.onMessage(chatJid, {
        id: ctx.message.message_id.toString(),
        chat_jid: chatJid,
        sender: senderId,
        sender_name: senderName,
        content: `${placeholder}${caption}`,
        timestamp,
        is_from_me: false,
      });
    };

    this.bot.on('message:photo', async (ctx) => {
      const chatJid = `tg:${ctx.chat.id}`;
      const group = this.opts.registeredGroups()[chatJid];
      const timestamp = new Date(ctx.message.date * 1000).toISOString();
      const senderName =
        ctx.from?.first_name ||
        ctx.from?.username ||
        ctx.from?.id?.toString() ||
        'Unknown';
      const caption = ctx.message.caption ? ` ${ctx.message.caption}` : '';
      const isGroup =
        ctx.chat.type === 'group' || ctx.chat.type === 'supergroup';
      this.opts.onChatMetadata(
        chatJid,
        timestamp,
        undefined,
        'telegram',
        isGroup,
      );
      if (!group) return;

      const photoSenderId = ctx.from?.id?.toString() || '';
      if (
        this.opts.allowedUserIds &&
        photoSenderId &&
        !this.opts.allowedUserIds.has(photoSenderId)
      ) {
        logger.warn(
          { chatJid, sender: senderName, userId: photoSenderId },
          'Telegram photo dropped: sender not in allowlist',
        );
        return;
      }

      // Pick the highest-resolution variant (last in the array)
      const photos = ctx.message.photo;
      const best = photos[photos.length - 1];
      let content = `[Photo]${caption}`;

      try {
        const file = await ctx.getFile();
        if (file.file_path) {
          const mediaDir = path.join(
            resolveGroupFolderPath(group.folder),
            'media',
          );
          fs.mkdirSync(mediaDir, { recursive: true });
          const filename = `${ctx.message.message_id}_${best.file_id.slice(-8)}.jpg`;
          const localPath = path.join(mediaDir, filename);
          const containerPath = `/workspace/group/media/${filename}`;
          const fileUrl = `https://api.telegram.org/file/bot${this.botToken}/${file.file_path}`;
          await downloadFile(fileUrl, localPath);
          content = `[Photo: ${containerPath}]${caption}`;
          logger.info({ chatJid, localPath }, 'Telegram photo downloaded');
        }
      } catch (err) {
        logger.warn(
          { err },
          'Failed to download Telegram photo, using placeholder',
        );
      }

      this.opts.onMessage(chatJid, {
        id: ctx.message.message_id.toString(),
        chat_jid: chatJid,
        sender: ctx.from?.id?.toString() || '',
        sender_name: senderName,
        content,
        timestamp,
        is_from_me: false,
      });
    });
    this.bot.on('message:video', (ctx) => storeNonText(ctx, '[Video]'));
    this.bot.on('message:voice', (ctx) => storeNonText(ctx, '[Voice message]'));
    this.bot.on('message:audio', (ctx) => storeNonText(ctx, '[Audio]'));
    this.bot.on('message:document', async (ctx) => {
      const doc = ctx.message.document;
      const name = doc?.file_name || 'file';
      const mime = doc?.mime_type || '';
      const ext = path.extname(name).toLowerCase();
      logger.info({ name, mime, ext, file_id: doc?.file_id }, 'Telegram document received');

      const isPdf = mime === 'application/pdf';
      const isExcel =
        mime ===
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        mime === 'application/vnd.ms-excel' ||
        ext === '.xlsx' ||
        ext === '.xls';
      const isWord =
        mime ===
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
        mime === 'application/msword' ||
        ext === '.docx';
      const isPowerPoint =
        mime ===
          'application/vnd.openxmlformats-officedocument.presentationml.presentation' ||
        mime === 'application/vnd.ms-powerpoint' ||
        ext === '.pptx';
      const isImage =
        mime.startsWith('image/') ||
        ['.png', '.jpg', '.jpeg', '.gif', '.webp'].includes(ext);
      const isText =
        mime.startsWith('text/') ||
        mime === 'application/json' ||
        mime === 'application/xml' ||
        mime === 'application/x-gedcom' ||
        ['.ged', '.txt', '.csv', '.json', '.xml', '.md', '.log'].includes(ext);

      // Image sent as document (e.g. PNG, JPEG): download to media/ like photo messages
      if (isImage && doc) {
        const chatJid = `tg:${ctx.chat.id}`;
        const group = this.opts.registeredGroups()[chatJid];
        if (!group) return;

        const imgSenderId = ctx.from?.id?.toString() || '';
        const timestamp = new Date(ctx.message.date * 1000).toISOString();
        const senderName =
          ctx.from?.first_name ||
          ctx.from?.username ||
          imgSenderId ||
          'Unknown';

        if (
          this.opts.allowedUserIds &&
          imgSenderId &&
          !this.opts.allowedUserIds.has(imgSenderId)
        ) {
          logger.warn(
            { chatJid, sender: senderName, userId: imgSenderId },
            'Telegram image dropped: sender not in allowlist',
          );
          return;
        }

        const caption = ctx.message.caption ? `\n${ctx.message.caption}` : '';
        const isGroup =
          ctx.chat.type === 'group' || ctx.chat.type === 'supergroup';
        this.opts.onChatMetadata(
          chatJid,
          timestamp,
          undefined,
          'telegram',
          isGroup,
        );

        let content = `[Photo]${caption}`;
        try {
          const file = await ctx.api.getFile(doc.file_id);
          if (file.file_path) {
            const mediaDir = path.join(
              resolveGroupFolderPath(group.folder),
              'media',
            );
            fs.mkdirSync(mediaDir, { recursive: true });
            const safeFilename = path
              .basename(name)
              .replace(/[^a-zA-Z0-9._-]/g, '_')
              .slice(0, 200);
            const localPath = path.join(mediaDir, safeFilename);
            const containerPath = `/workspace/group/media/${safeFilename}`;
            const fileUrl = `https://api.telegram.org/file/bot${this.botToken}/${file.file_path}`;
            await downloadFile(fileUrl, localPath);
            content = `[Photo: ${containerPath}]${caption}`;
            logger.info({ chatJid, localPath }, 'Telegram image (document) downloaded');
          }
        } catch (err) {
          logger.warn({ err }, 'Failed to download Telegram image document, using placeholder');
        }

        this.opts.onMessage(chatJid, {
          id: ctx.message.message_id.toString(),
          chat_jid: chatJid,
          sender: ctx.from?.id?.toString() || '',
          sender_name: senderName,
          content,
          timestamp,
          is_from_me: false,
        });
        return;
      }

      // PDF, Excel, Word, or PowerPoint: download and save for reader tools
      if ((isPdf || isExcel || isWord || isPowerPoint) && doc) {
        const chatJid = `tg:${ctx.chat.id}`;
        const group = this.opts.registeredGroups()[chatJid];
        if (!group) return;

        const docSenderId = ctx.from?.id?.toString() || '';
        const timestamp = new Date(ctx.message.date * 1000).toISOString();
        const senderName =
          ctx.from?.first_name ||
          ctx.from?.username ||
          docSenderId ||
          'Unknown';

        if (
          this.opts.allowedUserIds &&
          docSenderId &&
          !this.opts.allowedUserIds.has(docSenderId)
        ) {
          logger.warn(
            { chatJid, sender: senderName, userId: docSenderId },
            'Telegram document dropped: sender not in allowlist',
          );
          return;
        }

        const caption = ctx.message.caption || '';
        const isGroup =
          ctx.chat.type === 'group' || ctx.chat.type === 'supergroup';
        this.opts.onChatMetadata(
          chatJid,
          timestamp,
          undefined,
          'telegram',
          isGroup,
        );

        const safeDocFilename = path
          .basename(name)
          .replace(/[^a-zA-Z0-9._-]/g, '_')
          .slice(0, 200);
        let content = `[Document: ${safeDocFilename}]`;
        try {
          const file = await ctx.api.getFile(doc.file_id);
          if (file.file_path) {
            const downloadUrl = `https://api.telegram.org/file/bot${this.botToken}/${file.file_path}`;
            const groupDir = resolveGroupFolderPath(group.folder);
            const attachDir = path.join(groupDir, 'attachments');
            fs.mkdirSync(attachDir, { recursive: true });
            const filename = safeDocFilename;
            const filePath = path.join(attachDir, filename);
            await downloadFile(downloadUrl, filePath);
            const sizeKB = Math.round(fs.statSync(filePath).size / 1024);
            const ref = isPdf
              ? `[PDF: attachments/${filename} (${sizeKB}KB)]\nUse: pdf-reader extract attachments/${filename}`
              : `[Document: attachments/${filename} (${sizeKB}KB)]\nUse: doc-reader extract attachments/${filename}`;
            content = caption ? `${caption}\n\n${ref}` : ref;
            logger.info(
              { chatJid, filename },
              `Telegram ${isPdf ? 'PDF' : isExcel ? 'Excel' : isPowerPoint ? 'PowerPoint' : 'Word'} downloaded`,
            );
          }
        } catch (err) {
          logger.warn({ err, chatJid }, 'Document - download failed');
        }

        this.opts.onMessage(chatJid, {
          id: ctx.message.message_id.toString(),
          chat_jid: chatJid,
          sender: ctx.from?.id?.toString() || '',
          sender_name: senderName,
          content,
          timestamp,
          is_from_me: false,
        });
        return;
      }

      // Text-based files (GEDCOM, CSV, JSON, XML, etc.): download so the agent can read them
      if (isText && doc) {
        const chatJid = `tg:${ctx.chat.id}`;
        const group = this.opts.registeredGroups()[chatJid];
        if (!group) return;

        const txtSenderId = ctx.from?.id?.toString() || '';
        const timestamp = new Date(ctx.message.date * 1000).toISOString();
        const senderName =
          ctx.from?.first_name ||
          ctx.from?.username ||
          txtSenderId ||
          'Unknown';

        if (
          this.opts.allowedUserIds &&
          txtSenderId &&
          !this.opts.allowedUserIds.has(txtSenderId)
        ) {
          logger.warn(
            { chatJid, sender: senderName, userId: txtSenderId },
            'Telegram text document dropped: sender not in allowlist',
          );
          return;
        }

        const caption = ctx.message.caption ? `\n${ctx.message.caption}` : '';
        const isGroup =
          ctx.chat.type === 'group' || ctx.chat.type === 'supergroup';
        this.opts.onChatMetadata(
          chatJid,
          timestamp,
          undefined,
          'telegram',
          isGroup,
        );

        const safeTxtFilename = path
          .basename(name)
          .replace(/[^a-zA-Z0-9._-]/g, '_')
          .slice(0, 200);
        let content = `[Document: ${safeTxtFilename}]${caption}`;
        try {
          const file = await ctx.api.getFile(doc.file_id);
          if (file.file_path) {
            const groupDir = resolveGroupFolderPath(group.folder);
            const attachDir = path.join(groupDir, 'attachments');
            fs.mkdirSync(attachDir, { recursive: true });
            const safeFilename = safeTxtFilename;
            const filePath = path.join(attachDir, safeFilename);
            const fileUrl = `https://api.telegram.org/file/bot${this.botToken}/${file.file_path}`;
            await downloadFile(fileUrl, filePath);
            const sizeKB = Math.round(fs.statSync(filePath).size / 1024);
            const ref = `[File: attachments/${safeFilename} (${sizeKB}KB)]\nUse: cat attachments/${safeFilename}`;
            content = caption ? `${caption}\n\n${ref}` : ref;
            logger.info({ chatJid, filePath }, 'Telegram text document downloaded');
          }
        } catch (err) {
          logger.warn({ err }, 'Failed to download Telegram text document, using placeholder');
        }

        this.opts.onMessage(chatJid, {
          id: ctx.message.message_id.toString(),
          chat_jid: chatJid,
          sender: ctx.from?.id?.toString() || '',
          sender_name: senderName,
          content,
          timestamp,
          is_from_me: false,
        });
        return;
      }

      // Unknown document type: still download so the agent can inspect it
      if (doc) {
        const chatJid = `tg:${ctx.chat.id}`;
        const group = this.opts.registeredGroups()[chatJid];
        if (!group) {
          storeNonText(ctx, `[Document: ${name}]`);
          return;
        }

        const unknownSenderId = ctx.from?.id?.toString() || '';
        const timestamp = new Date(ctx.message.date * 1000).toISOString();
        const senderName =
          ctx.from?.first_name ||
          ctx.from?.username ||
          unknownSenderId ||
          'Unknown';

        if (
          this.opts.allowedUserIds &&
          unknownSenderId &&
          !this.opts.allowedUserIds.has(unknownSenderId)
        ) {
          logger.warn(
            { chatJid, sender: senderName, userId: unknownSenderId },
            'Telegram unknown document dropped: sender not in allowlist',
          );
          return;
        }

        const caption = ctx.message.caption ? `\n${ctx.message.caption}` : '';
        const isGroup =
          ctx.chat.type === 'group' || ctx.chat.type === 'supergroup';
        this.opts.onChatMetadata(
          chatJid,
          timestamp,
          undefined,
          'telegram',
          isGroup,
        );

        const safeUnknownFilename = path
          .basename(name)
          .replace(/[^a-zA-Z0-9._-]/g, '_')
          .slice(0, 200);
        let content = `[Document: ${safeUnknownFilename}]${caption}`;
        try {
          const file = await ctx.api.getFile(doc.file_id);
          if (file.file_path) {
            const groupDir = resolveGroupFolderPath(group.folder);
            const attachDir = path.join(groupDir, 'attachments');
            fs.mkdirSync(attachDir, { recursive: true });
            const safeFilename = safeUnknownFilename;
            const filePath = path.join(attachDir, safeFilename);
            const fileUrl = `https://api.telegram.org/file/bot${this.botToken}/${file.file_path}`;
            await downloadFile(fileUrl, filePath);
            const sizeKB = Math.round(fs.statSync(filePath).size / 1024);
            const ref = `[File: attachments/${safeFilename} (${sizeKB}KB)]\nUse: cat attachments/${safeFilename}`;
            content = caption ? `${caption}\n\n${ref}` : ref;
            logger.info({ chatJid, filePath, mime, ext }, 'Telegram unknown document downloaded');
          }
        } catch (err) {
          logger.warn({ err, name, mime }, 'Failed to download Telegram unknown document');
        }

        this.opts.onMessage(chatJid, {
          id: ctx.message.message_id.toString(),
          chat_jid: chatJid,
          sender: ctx.from?.id?.toString() || '',
          sender_name: senderName,
          content,
          timestamp,
          is_from_me: false,
        });
        return;
      }

      const safeDocName = path.basename(name).replace(/[^a-zA-Z0-9._-]/g, '_').slice(0, 200);
      storeNonText(ctx, `[Document: ${safeDocName}]`);
    });
    this.bot.on('message:sticker', (ctx) => {
      const emoji = ctx.message.sticker?.emoji || '';
      storeNonText(ctx, `[Sticker ${emoji}]`);
    });
    this.bot.on('message:location', (ctx) => storeNonText(ctx, '[Location]'));
    this.bot.on('message:contact', (ctx) => storeNonText(ctx, '[Contact]'));

    // Handle errors gracefully
    this.bot.catch((err) => {
      logger.error({ err: err.message }, 'Telegram bot error');
    });

    // Start polling — returns a Promise that resolves when started
    return new Promise<void>((resolve) => {
      this.bot!.start({
        onStart: (botInfo) => {
          logger.info(
            { username: botInfo.username, id: botInfo.id },
            'Telegram bot connected',
          );
          console.log(`\n  Telegram bot: @${botInfo.username}`);
          console.log(
            `  Send /chatid to the bot to get a chat's registration ID\n`,
          );
          resolve();
        },
      });
    });
  }

  async sendMessage(jid: string, text: string): Promise<void> {
    if (!this.bot) {
      logger.warn('Telegram bot not initialized');
      return;
    }

    try {
      const numericId = jid.replace(/^tg:/, '');

      // Telegram has a 4096 character limit per message — split if needed
      const MAX_LENGTH = 4096;
      if (text.length <= MAX_LENGTH) {
        await sendTelegramMessage(this.bot.api, numericId, text);
      } else {
        for (let i = 0; i < text.length; i += MAX_LENGTH) {
          await sendTelegramMessage(
            this.bot.api,
            numericId,
            text.slice(i, i + MAX_LENGTH),
          );
        }
      }
      logger.info({ jid, length: text.length }, 'Telegram message sent');
    } catch (err) {
      logger.error({ jid, err }, 'Failed to send Telegram message');
    }
  }

  isConnected(): boolean {
    return this.bot !== null;
  }

  ownsJid(jid: string): boolean {
    return jid.startsWith('tg:');
  }

  async disconnect(): Promise<void> {
    if (this.bot) {
      this.bot.stop();
      this.bot = null;
      logger.info('Telegram bot stopped');
    }
  }

  async setTyping(jid: string, isTyping: boolean): Promise<void> {
    if (!this.bot || !isTyping) return;
    try {
      const numericId = jid.replace(/^tg:/, '');
      await this.bot.api.sendChatAction(numericId, 'typing');
    } catch (err) {
      logger.debug({ jid, err }, 'Failed to send Telegram typing indicator');
    }
  }
}

// Bot pool for agent teams: send-only Api instances (no polling)
const poolApis: Api[] = [];
const senderBotMap = new Map<string, number>();
let nextPoolIndex = 0;

export async function initBotPool(tokens: string[]): Promise<void> {
  for (const token of tokens) {
    try {
      const api = new Api(token);
      const me = await api.getMe();
      poolApis.push(api);
      logger.info(
        { username: me.username, id: me.id, poolSize: poolApis.length },
        'Pool bot initialized',
      );
    } catch (err) {
      logger.error({ err }, 'Failed to initialize pool bot');
    }
  }
  if (poolApis.length > 0) {
    logger.info({ count: poolApis.length }, 'Telegram bot pool ready');
  }
}

export async function sendPoolMessage(
  chatId: string,
  text: string,
  sender: string,
  groupFolder: string,
): Promise<void> {
  if (poolApis.length === 0) {
    logger.warn({ sender }, 'No pool bots available, dropping pool message');
    return;
  }

  const key = `${groupFolder}:${sender}`;
  let idx = senderBotMap.get(key);
  if (idx === undefined) {
    idx = nextPoolIndex % poolApis.length;
    nextPoolIndex++;
    senderBotMap.set(key, idx);
    try {
      const safeName = sender.replace(/[^\w\s\-().]/gu, '').slice(0, 64);
      await poolApis[idx].setMyName(safeName);
      await new Promise((r) => setTimeout(r, 2000));
      logger.info(
        { sender, groupFolder, poolIndex: idx },
        'Assigned and renamed pool bot',
      );
    } catch (err) {
      logger.warn(
        { sender, err },
        'Failed to rename pool bot (sending anyway)',
      );
    }
  }

  const api = poolApis[idx];
  try {
    const numericId = chatId.replace(/^tg:/, '');
    const MAX_LENGTH = 4096;
    if (text.length <= MAX_LENGTH) {
      await api.sendMessage(numericId, text);
    } else {
      for (let i = 0; i < text.length; i += MAX_LENGTH) {
        await api.sendMessage(numericId, text.slice(i, i + MAX_LENGTH));
      }
    }
    logger.info(
      { chatId, sender, poolIndex: idx, length: text.length },
      'Pool message sent',
    );
  } catch (err) {
    logger.error({ chatId, sender, err }, 'Failed to send pool message');
  }
}

registerChannel('telegram', (opts: ChannelOpts) => {
  const envVars = readEnvFile([
    'TELEGRAM_BOT_TOKEN',
    'TELEGRAM_ALLOWED_USER_IDS',
  ]);
  const token =
    process.env.TELEGRAM_BOT_TOKEN || envVars.TELEGRAM_BOT_TOKEN || '';
  if (!token) {
    logger.warn('Telegram: TELEGRAM_BOT_TOKEN not set');
    return null;
  }

  const rawAllowed =
    process.env.TELEGRAM_ALLOWED_USER_IDS ||
    envVars.TELEGRAM_ALLOWED_USER_IDS ||
    '';
  const allowedUserIds = rawAllowed
    ? new Set(
        rawAllowed
          .split(',')
          .map((id) => id.trim())
          .filter(Boolean),
      )
    : undefined;

  if (allowedUserIds) {
    logger.info(
      { count: allowedUserIds.size },
      'Telegram: sender allowlist active',
    );
  }

  return new TelegramChannel(token, { ...opts, allowedUserIds });
});
