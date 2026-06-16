import type { Route } from "./+types/home";
import { requireAuth } from "~/lib/auth";
import { getAllStorages, getPublicStorages, initDatabase } from "~/lib/storage";
import { useState, useEffect, useCallback } from "react";
import { FilePreview } from "~/components/FilePreview";
import { getFileType, isPreviewable } from "~/lib/file-utils";

export function meta({ data }: Route.MetaArgs) {
  const title = data?.siteTitle || "CList";
  return [
    { title: `${title} - 存储聚合` },
    { name: "description", content: "S3 兼容存储聚合服务" },
  ];
}

export async function loader({ request, context }: Route.LoaderArgs) {
  const db = context.cloudflare.env.DB;
  const siteTitle = context.cloudflare.env.SITE_TITLE || "CList";
  const siteAnnouncement = context.cloudflare.env.SITE_ANNOUNCEMENT || "";
  const chunkSizeMB = parseInt(context.cloudflare.env.CHUNK_SIZE_MB || "50", 10);
  const webdavEnabled = (context.cloudflare.env.WEBDAV_ENABLED as string) === "true";

  if (!db) {
    console.error("D1 Database not bound");
    return { isAdmin: false, storages: [], siteTitle, siteAnnouncement, chunkSizeMB, webdavEnabled: false };
  }

  await initDatabase(db);

  const { isAdmin } = await requireAuth(request, db);

  const storages = isAdmin
    ? await getAllStorages(db)
    : await getPublicStorages(db);

  return {
    isAdmin,
    siteTitle,
    siteAnnouncement,
    chunkSizeMB,
    webdavEnabled,
    storages: storages.map((s) => ({
      id: s.id,
      name: s.name,
      type: s.type,
      endpoint: s.endpoint,
      region: s.region,
      accessKeyId: s.accessKeyId,
      bucket: s.bucket,
      basePath: s.basePath,
      config: isAdmin ? s.config : undefined,
      isPublic: s.isPublic,
      guestList: s.guestList,
      guestDownload: s.guestDownload,
      guestUpload: s.guestUpload,
    })),
  };
}

interface S3Object {
  key: string;
  name: string;
  size: number;
  lastModified: string;
  isDirectory: boolean;
}

interface StorageInfo {
  id: number;
  name: string;
  type?: string;
  endpoint?: string;
  region?: string;
  accessKeyId?: string;
  bucket?: string;
  basePath?: string;
  config?: Record<string, any>;
  isPublic: boolean;
  guestList: boolean;
  guestDownload: boolean;
  guestUpload: boolean;
}

type ConfigField = {
  key: string;
  label: string;
  type: "text" | "password" | "textarea" | "select" | "boolean";
  required?: boolean;
  placeholder?: string;
  options?: Array<{ value: string; label: string }>;
  defaultValue?: string | number | boolean;
  show?: (values: Record<string, any>) => boolean;
  help?: string;
};

const driveConfigMap: Record<string, { name: string; supportsMultipart: boolean; fields: ConfigField[] }> = {
  onedrive: {
    name: "OneDrive",
    supportsMultipart: true,
    fields: [
      {
        key: "region",
        label: "区域",
        type: "select",
        required: true,
        options: [
          { value: "global", label: "全球版" },
          { value: "cn", label: "中国版（世纪互联）" },
          { value: "us", label: "美国政府版" },
          { value: "de", label: "德国版" },
        ],
        defaultValue: "global",
      },
      { key: "refresh_token", label: "刷新令牌", type: "textarea", required: true, placeholder: "Microsoft OAuth 刷新令牌" },
      { key: "use_online_api", label: "使用在线API", type: "boolean", defaultValue: true },
      {
        key: "api_address",
        label: "在线API地址",
        type: "text",
        defaultValue: "https://api.oplist.org/onedrive/renewapi",
        placeholder: "自建刷新接口地址",
        show: (values) => values.use_online_api === true,
      },
      {
        key: "client_id",
        label: "客户端ID",
        type: "text",
        placeholder: "本地客户端ID",
        show: (values) => values.use_online_api !== true,
      },
      {
        key: "client_secret",
        label: "客户端密钥",
        type: "password",
        placeholder: "本地客户端密钥",
        show: (values) => values.use_online_api !== true,
      },
      {
        key: "redirect_uri",
        label: "重定向URI",
        type: "text",
        placeholder: "https://api.oplist.org/onedrive/callback",
        defaultValue: "https://api.oplist.org/onedrive/callback",
        show: (values) => values.use_online_api !== true,
      },
      { key: "is_sharepoint", label: "SharePoint 模式", type: "boolean", defaultValue: false },
      {
        key: "site_id",
        label: "SharePoint 站点ID",
        type: "text",
        placeholder: "SharePoint 站点ID",
        show: (values) => values.is_sharepoint === true,
      },
      { key: "root_folder_path", label: "根文件夹路径", type: "text", defaultValue: "/" },
      { key: "chunk_size", label: "分块大小 (MB)", type: "text", defaultValue: "5" },
      { key: "custom_host", label: "自定义下载主机", type: "text", placeholder: "可选：自定义下载域名" },
    ],
  },
  gdrive: {
    name: "Google Drive",
    supportsMultipart: true,
    fields: [
      { key: "refresh_token", label: "刷新令牌", type: "textarea", required: true, placeholder: "Google OAuth 刷新令牌" },
      { key: "use_online_api", label: "使用在线API", type: "boolean", defaultValue: true },
      {
        key: "api_address",
        label: "在线API地址",
        type: "text",
        defaultValue: "https://api.oplist.org/googleui/renewapi",
        placeholder: "自建刷新接口地址",
        show: (values) => values.use_online_api === true,
      },
      {
        key: "client_id",
        label: "客户端ID",
        type: "text",
        placeholder: "本地客户端ID",
        show: (values) => values.use_online_api !== true,
      },
      {
        key: "client_secret",
        label: "客户端密钥",
        type: "password",
        placeholder: "本地客户端密钥",
        show: (values) => values.use_online_api !== true,
      },
      { key: "root_folder_id", label: "根目录ID", type: "text", defaultValue: "root", placeholder: "默认 root" },
      { key: "order_by", label: "排序字段", type: "text", defaultValue: "folder,name,modifiedTime", placeholder: "folder,name,modifiedTime" },
      {
        key: "order_direction",
        label: "排序方向",
        type: "select",
        options: [
          { value: "asc", label: "升序" },
          { value: "desc", label: "降序" },
        ],
        defaultValue: "asc",
      },
      { key: "chunk_size", label: "分块大小 (MB)", type: "text", defaultValue: "5" },
    ],
  },
  alicloud: {
    name: "阿里云盘",
    supportsMultipart: true,
    fields: [
      {
        key: "drive_type",
        label: "驱动类型",
        type: "select",
        required: true,
        options: [
          { value: "resource", label: "资源库" },
          { value: "backup", label: "备份盘" },
          { value: "default", label: "默认" },
        ],
        defaultValue: "resource",
      },
      { key: "refresh_token", label: "刷新令牌", type: "textarea", required: true },
      { key: "root_folder_id", label: "根目录ID", type: "text", defaultValue: "root" },
      {
        key: "order_by",
        label: "排序方式",
        type: "select",
        options: [
          { value: "name", label: "文件名" },
          { value: "size", label: "文件大小" },
          { value: "updated_at", label: "修改时间" },
          { value: "created_at", label: "创建时间" },
        ],
        defaultValue: "name",
      },
      {
        key: "order_direction",
        label: "排序方向",
        type: "select",
        options: [
          { value: "ASC", label: "升序" },
          { value: "DESC", label: "降序" },
        ],
        defaultValue: "ASC",
      },
      { key: "use_online_api", label: "使用在线API", type: "boolean", defaultValue: true },
      {
        key: "api_address",
        label: "在线API地址",
        type: "text",
        defaultValue: "https://api.oplist.org/alicloud/renewapi",
        placeholder: "自建刷新接口地址",
        show: (values) => values.use_online_api === true,
      },
      {
        key: "client_id",
        label: "客户端ID",
        type: "text",
        placeholder: "本地客户端ID",
        show: (values) => values.use_online_api !== true,
      },
      {
        key: "client_secret",
        label: "客户端密钥",
        type: "password",
        placeholder: "本地客户端密钥",
        show: (values) => values.use_online_api !== true,
      },
      {
        key: "remove_way",
        label: "删除方式",
        type: "select",
        options: [
          { value: "trash", label: "移到回收站" },
          { value: "delete", label: "直接删除" },
        ],
        defaultValue: "trash",
      },
      { key: "rapid_upload", label: "秒传", type: "boolean", defaultValue: false },
      { key: "internal_upload", label: "内网上传", type: "boolean", defaultValue: false },
      {
        key: "livp_download_format",
        label: "LIVP 下载格式",
        type: "select",
        options: [
          { value: "jpeg", label: "JPEG" },
          { value: "mov", label: "MOV" },
        ],
        defaultValue: "jpeg",
      },
      {
        key: "alipan_type",
        label: "云盘类型",
        type: "select",
        options: [
          { value: "default", label: "默认" },
          { value: "alipanTV", label: "阿里云盘TV" },
        ],
        defaultValue: "default",
      },
    ],
  },
  baiduyun: {
    name: "百度网盘",
    supportsMultipart: false,
    fields: [
      { key: "refresh_token", label: "刷新令牌", type: "textarea", required: true },
      { key: "root_path", label: "根目录路径", type: "text", defaultValue: "/" },
      {
        key: "order_by",
        label: "排序方式",
        type: "select",
        options: [
          { value: "name", label: "文件名" },
          { value: "time", label: "修改时间" },
          { value: "size", label: "文件大小" },
        ],
        defaultValue: "name",
      },
      {
        key: "order_direction",
        label: "排序方向",
        type: "select",
        options: [
          { value: "asc", label: "升序" },
          { value: "desc", label: "降序" },
        ],
        defaultValue: "asc",
      },
      { key: "use_online_api", label: "使用在线API", type: "boolean", defaultValue: true },
      {
        key: "api_address",
        label: "在线API地址",
        type: "text",
        defaultValue: "https://api.oplist.org/baiduyun/renewapi",
        placeholder: "自建刷新接口地址",
        show: (values) => values.use_online_api === true,
      },
      {
        key: "client_id",
        label: "客户端ID",
        type: "text",
        placeholder: "本地客户端ID",
        show: (values) => values.use_online_api !== true,
      },
      {
        key: "client_secret",
        label: "客户端密钥",
        type: "password",
        placeholder: "本地客户端密钥",
        show: (values) => values.use_online_api !== true,
      },
    ],
  },
};

function supportsMultipart(type?: string): boolean {
  if (!type) {
    return true;
  }
  if (type === "webdev") {
    return false;
  }
  if (type === "s3") {
    return true;
  }
  const config = driveConfigMap[type];
  if (config) {
    return config.supportsMultipart;
  }
  return false;
}

interface AuditLog {
  id: number;
  action: string;
  storageId: number | null;
  path: string | null;
  userType: "guest" | "admin" | "share";
  ip: string | null;
  userAgent: string | null;
  detail: string | null;
  createdAt: string;
}


function formatBytes(bytes: number): string {
  if (bytes === 0) return "-";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB", "TB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

function formatSpeed(bytesPerSecond: number): string {
  if (bytesPerSecond === 0) return "0 B/s";
  const k = 1024;
  const sizes = ["B/s", "KB/s", "MB/s", "GB/s"];
  const i = Math.floor(Math.log(bytesPerSecond) / Math.log(k));
  return parseFloat((bytesPerSecond / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

function formatDate(dateStr: string): string {
  if (!dateStr) return "-";
  const date = new Date(dateStr);
  return date.toLocaleString("zh-CN");
}

function StatsIcon({ className = "h-4 w-4" }: { className?: string }) {
  return (
    <svg viewBox="0 0 24 24" className={className} fill="none" aria-hidden="true">
      <path
        d="M5 19h14"
        stroke="currentColor"
        strokeWidth="1.7"
        strokeLinecap="round"
        opacity="0.75"
      />
      <path
        d="M7 16v-3.5M12 16V8M17 16v-5"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinecap="round"
      />
      <path
        d="M5.5 10.5 9 7l3.2 3.2 5.7-5.7"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
      <path
        d="M16.5 4.5h1.9v1.9"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function LoginModal({ onLogin, onClose }: { onLogin: () => void; onClose: () => void }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError("");

    try {
      const res = await fetch("/api/storages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "login", username, password }),
      });

      if (res.ok) {
        onLogin();
      } else {
        const data = (await res.json()) as { error?: string };
        setError(data.error || "登录失败");
      }
    } catch {
      setError("网络错误");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black/50 dark:bg-black/70 backdrop-blur-sm flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700/50 w-full max-w-sm rounded-xl shadow-2xl clist-modal-enter" onClick={e => e.stopPropagation()}>
        <div className="px-5 py-4 border-b border-zinc-200 dark:border-zinc-700/50 flex items-center justify-between">
          <span className="text-zinc-900 dark:text-zinc-100 font-medium text-sm">管理员登录</span>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>
        </div>
        <form onSubmit={handleSubmit} className="p-5 space-y-4">
          <div>
            <label className="block text-xs text-zinc-500 mb-1.5 font-medium">用户名</label>
            <input
              type="text"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              className="w-full bg-zinc-50 dark:bg-zinc-800/80 border border-zinc-200 dark:border-zinc-700 px-3 py-2.5 text-zinc-900 dark:text-zinc-100 text-sm rounded-lg focus:border-blue-500 focus:ring-1 focus:ring-blue-500/20 focus:outline-none transition"
              required
            />
          </div>
          <div>
            <label className="block text-xs text-zinc-500 mb-1.5 font-medium">密码</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full bg-zinc-50 dark:bg-zinc-800/80 border border-zinc-200 dark:border-zinc-700 px-3 py-2.5 text-zinc-900 dark:text-zinc-100 text-sm rounded-lg focus:border-blue-500 focus:ring-1 focus:ring-blue-500/20 focus:outline-none transition"
              required
            />
          </div>
          {error && <div className="text-red-500 dark:text-red-400 text-xs bg-red-50 dark:bg-red-900/20 px-3 py-2 rounded-lg">{error}</div>}
          <div className="flex gap-2 pt-1">
            <button
              type="button"
              onClick={onClose}
              className="flex-1 py-2.5 px-4 border border-zinc-200 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400 hover:text-zinc-900 dark:hover:text-zinc-100 hover:bg-zinc-50 dark:hover:bg-zinc-800 text-sm font-medium transition rounded-lg"
            >
              取消
            </button>
            <button
              type="submit"
              disabled={loading}
              className="flex-1 py-2.5 px-4 bg-blue-600 hover:bg-blue-500 text-white text-sm font-medium disabled:opacity-50 transition rounded-lg shadow-sm"
            >
              {loading ? "登录中..." : "登录"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

function StorageModal({
  storage,
  onSave,
  onCancel,
}: {
  storage?: StorageInfo;
  onSave: () => void;
  onCancel: () => void;
}) {
  const initConfig = (type: string, existing?: Record<string, any>) => {
    const fields = driveConfigMap[type]?.fields || [];
    const base = { ...(existing || {}) };
    if (base.api_address === undefined && base.api_url_address !== undefined) {
      base.api_address = base.api_url_address;
    }
    for (const field of fields) {
      if (base[field.key] === undefined && field.defaultValue !== undefined) {
        base[field.key] = field.defaultValue;
      }
    }
    const hasLocalClient = Boolean(String(base.client_id || "").trim() && String(base.client_secret || "").trim());
    if (fields.some((field) => field.key === "use_online_api") && !hasLocalClient) {
      base.use_online_api = true;
    }
    return base;
  };

  const [formData, setFormData] = useState({
    name: storage?.name || "",
    type: storage?.type || "s3",
    endpoint: storage?.endpoint || "",
    region: storage?.region || "auto",
    accessKeyId: storage?.accessKeyId || "",
    secretAccessKey: "",
    bucket: storage?.bucket || "",
    basePath: storage?.basePath || "",
    config: initConfig(storage?.type || "s3", storage?.config),
    isPublic: storage?.isPublic ?? false,
    guestList: storage?.guestList ?? false,
    guestDownload: storage?.guestDownload ?? false,
    guestUpload: storage?.guestUpload ?? false,
  });
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const driveConfig = driveConfigMap[formData.type || ""];
  const isS3 = formData.type === "s3";
  const isWebdav = formData.type === "webdev";

  const handleTypeChange = (nextType: string) => {
    setFormData({
      ...formData,
      type: nextType,
      endpoint: nextType === "s3" || nextType === "webdev" ? formData.endpoint : "",
      region: nextType === "s3" ? formData.region : "auto",
      accessKeyId: nextType === "s3" || nextType === "webdev" ? formData.accessKeyId : "",
      secretAccessKey: "",
      bucket: nextType === "s3" ? formData.bucket : "",
      basePath: nextType === "s3" || nextType === "webdev" ? formData.basePath : "",
      config: initConfig(nextType, {}),
    });
  };

  const updateConfigValue = (key: string, value: string | number | boolean) => {
    setFormData({
      ...formData,
      config: { ...(formData.config || {}), [key]: value },
    });
  };

  const renderConfigField = (field: ConfigField) => {
    const values = formData.config || {};
    if (field.show && !field.show(values)) {
      return null;
    }

    const commonClasses = "w-full bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 px-3 py-2 text-zinc-900 dark:text-zinc-100 font-mono text-sm focus:border-blue-500 focus:outline-none rounded";
    const value = values[field.key] ?? "";

    if (field.type === "boolean") {
      return (
        <label key={field.key} className="flex items-center gap-2 cursor-pointer">
          <input
            type="checkbox"
            checked={Boolean(value)}
            onChange={(e) => updateConfigValue(field.key, e.target.checked)}
            className="w-4 h-4 bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded"
          />
          <span className="text-sm text-zinc-700 dark:text-zinc-300 font-mono">{field.label}</span>
          {field.help && <span className="text-xs text-zinc-500">{field.help}</span>}
        </label>
      );
    }

    if (field.type === "select") {
      return (
        <div key={field.key}>
          <label className="block text-xs text-zinc-500 mb-1 font-mono">{field.label}{field.required ? " *" : ""}</label>
          <select
            value={String(value)}
            onChange={(e) => updateConfigValue(field.key, e.target.value)}
            className={commonClasses}
            required={field.required}
          >
            {(field.options || []).map((opt) => (
              <option key={opt.value} value={opt.value}>{opt.label}</option>
            ))}
          </select>
        </div>
      );
    }

    if (field.type === "textarea") {
      return (
        <div key={field.key}>
          <label className="block text-xs text-zinc-500 mb-1 font-mono">{field.label}{field.required ? " *" : ""}</label>
          <textarea
            value={String(value)}
            onChange={(e) => updateConfigValue(field.key, e.target.value)}
            className={`${commonClasses} h-24`}
            placeholder={field.placeholder || ""}
            required={field.required}
          />
        </div>
      );
    }

    return (
      <div key={field.key}>
        <label className="block text-xs text-zinc-500 mb-1 font-mono">{field.label}{field.required ? " *" : ""}</label>
        <input
          type={field.type}
          value={String(value)}
          onChange={(e) => updateConfigValue(field.key, e.target.value)}
          className={commonClasses}
          placeholder={field.placeholder || ""}
          required={field.required}
        />
      </div>
    );
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError("");

    try {
      const method = storage ? "PUT" : "POST";
      const configToSend = { ...(formData.config || {}) };
      if (configToSend.api_address && !configToSend.api_url_address) {
        configToSend.api_url_address = configToSend.api_address;
      }
      if (driveConfig) {
        for (const field of driveConfig.fields) {
          if (field.type === "password" && !configToSend[field.key]) {
            delete configToSend[field.key];
          }
        }
      }
      const body = storage
        ? { id: storage.id, ...formData, config: configToSend }
        : { ...formData, config: configToSend };

      if (storage && !formData.secretAccessKey && (isS3 || isWebdav)) {
        delete (body as Record<string, unknown>).secretAccessKey;
      }

      const res = await fetch("/api/storages", {
        method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });

      if (res.ok) {
        onSave();
      } else {
        const data = (await res.json()) as { error?: string };
        setError(data.error || "保存失败");
      }
    } catch {
      setError("网络错误");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black/50 dark:bg-black/70 backdrop-blur-sm flex items-center justify-center z-50 p-4" onClick={onCancel}>
      <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700/50 w-full max-w-lg max-h-[90vh] overflow-y-auto rounded-xl shadow-2xl clist-modal-enter" onClick={e => e.stopPropagation()}>
        <div className="px-5 py-4 border-b border-zinc-200 dark:border-zinc-700/50 flex items-center justify-between sticky top-0 bg-white dark:bg-zinc-900 rounded-t-xl z-10">
          <span className="text-zinc-900 dark:text-zinc-100 font-medium text-sm">{storage ? "编辑存储" : "添加存储"}</span>
          <button onClick={onCancel} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>
        </div>
        <form onSubmit={handleSubmit} className="p-5 space-y-4">
          <div className="grid grid-cols-2 gap-4">
            <div className="col-span-2">
              <label className="block text-xs text-zinc-500 mb-1 font-mono">名称 *</label>
              <input
                type="text"
                value={formData.name}
                onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                className="w-full bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 px-3 py-2 text-zinc-900 dark:text-zinc-100 font-mono text-sm focus:border-blue-500 focus:outline-none rounded"
                placeholder="My Storage"
                required
              />
            </div>
            <div className="col-span-2">
              <label className="block text-xs text-zinc-500 mb-1 font-mono">存储类型 *</label>
              <select
                value={formData.type}
                onChange={(e) => handleTypeChange(e.target.value)}
                className="w-full bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 px-3 py-2 text-zinc-900 dark:text-zinc-100 font-mono text-sm focus:border-blue-500 focus:outline-none rounded"
                required
              >
                <option value="s3">S3 兼容服务</option>
                <option value="webdev">WebDAV</option>
                <option value="onedrive">OneDrive</option>
                <option value="gdrive">Google Drive</option>
                <option value="alicloud">阿里云盘</option>
                <option value="baiduyun">百度网盘</option>
              </select>
            </div>
            {(isS3 || isWebdav) && (
              <div className="col-span-2">
                <label className="block text-xs text-zinc-500 mb-1 font-mono">
                  {isWebdav ? "WebDAV 服务器地址" : "Endpoint"} *
                </label>
                <input
                  type="url"
                  value={formData.endpoint}
                  onChange={(e) => setFormData({ ...formData, endpoint: e.target.value })}
                  className="w-full bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 px-3 py-2 text-zinc-900 dark:text-zinc-100 font-mono text-sm focus:border-blue-500 focus:outline-none rounded"
                  placeholder={isWebdav ? "https://example.com/webdav" : "https://s3.us-east-1.amazonaws.com"}
                  required
                />
              </div>
            )}
            {isS3 && (
              <div>
                <label className="block text-xs text-zinc-500 mb-1 font-mono">Region</label>
                <input
                  type="text"
                  value={formData.region}
                  onChange={(e) => setFormData({ ...formData, region: e.target.value })}
                  className="w-full bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 px-3 py-2 text-zinc-900 dark:text-zinc-100 font-mono text-sm focus:border-blue-500 focus:outline-none rounded"
                  placeholder="auto"
                />
              </div>
            )}
            {isS3 && (
              <div>
                <label className="block text-xs text-zinc-500 mb-1 font-mono">Bucket *</label>
                <input
                  type="text"
                  value={formData.bucket}
                  onChange={(e) => setFormData({ ...formData, bucket: e.target.value })}
                  className="w-full bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 px-3 py-2 text-zinc-900 dark:text-zinc-100 font-mono text-sm focus:border-blue-500 focus:outline-none rounded"
                  placeholder="my-bucket"
                  required={isS3}
                />
              </div>
            )}
            {(isS3 || isWebdav) && (
              <>
                <div>
                  <label className="block text-xs text-zinc-500 mb-1 font-mono">
                    {isWebdav ? "用户名" : "Access Key"} *
                  </label>
                  <input
                    type="text"
                    value={formData.accessKeyId}
                    onChange={(e) => setFormData({ ...formData, accessKeyId: e.target.value })}
                    className="w-full bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 px-3 py-2 text-zinc-900 dark:text-zinc-100 font-mono text-sm focus:border-blue-500 focus:outline-none rounded"
                    required={!storage && (isS3 || isWebdav)}
                  />
                </div>
                <div>
                  <label className="block text-xs text-zinc-500 mb-1 font-mono">
                    {isWebdav ? "密码" : "Secret Key"} {storage && "(留空保持)"}
                  </label>
                  <input
                    type="password"
                    value={formData.secretAccessKey}
                    onChange={(e) => setFormData({ ...formData, secretAccessKey: e.target.value })}
                    className="w-full bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 px-3 py-2 text-zinc-900 dark:text-zinc-100 font-mono text-sm focus:border-blue-500 focus:outline-none rounded"
                    required={!storage && (isS3 || isWebdav)}
                  />
                </div>
                <div className="col-span-2">
                  <label className="block text-xs text-zinc-500 mb-1 font-mono">根路径</label>
                  <input
                    type="text"
                    value={formData.basePath}
                    onChange={(e) => setFormData({ ...formData, basePath: e.target.value })}
                    className="w-full bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 px-3 py-2 text-zinc-900 dark:text-zinc-100 font-mono text-sm focus:border-blue-500 focus:outline-none rounded"
                    placeholder="/path/to/folder"
                  />
                </div>
              </>
            )}
            {driveConfig && (
              <div className="col-span-2 border-t border-zinc-200 dark:border-zinc-700 pt-3 mt-1">
                <div className="text-xs text-zinc-500 mb-2 font-mono">驱动配置 - {driveConfig.name}</div>
                <div className="space-y-3">
                  {driveConfig.fields.map(renderConfigField)}
                </div>
              </div>
            )}
            <div className="col-span-2">
              <label className="flex items-center gap-2 cursor-pointer">
                <input
                  type="checkbox"
                  checked={formData.isPublic}
                  onChange={(e) => {
                    const checked = e.target.checked;
                    setFormData({
                      ...formData,
                      isPublic: checked,
                      guestList: checked,
                      guestDownload: checked,
                    });
                  }}
                  className="w-4 h-4 bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded"
                />
                <span className="text-sm text-zinc-700 dark:text-zinc-300 font-mono">公开访问</span>
                <span className="text-xs text-zinc-500">(快速开启浏览和下载)</span>
              </label>
            </div>
            <div className="col-span-2 border-t border-zinc-200 dark:border-zinc-700 pt-3 mt-1">
              <div className="text-xs text-zinc-500 mb-2 font-mono">游客权限设置</div>
              <div className="space-y-2">
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={formData.guestList}
                    onChange={(e) => setFormData({ ...formData, guestList: e.target.checked })}
                    className="w-4 h-4 bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded"
                  />
                  <span className="text-sm text-zinc-700 dark:text-zinc-300 font-mono">允许浏览</span>
                  <span className="text-xs text-zinc-500">(查看文件列表)</span>
                </label>
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={formData.guestDownload}
                    onChange={(e) => setFormData({ ...formData, guestDownload: e.target.checked })}
                    className="w-4 h-4 bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded"
                  />
                  <span className="text-sm text-zinc-700 dark:text-zinc-300 font-mono">允许下载</span>
                  <span className="text-xs text-zinc-500">(下载和预览文件)</span>
                </label>
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={formData.guestUpload}
                    onChange={(e) => setFormData({ ...formData, guestUpload: e.target.checked })}
                    className="w-4 h-4 bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded"
                  />
                  <span className="text-sm text-zinc-700 dark:text-zinc-300 font-mono">允许上传</span>
                  <span className="text-xs text-zinc-500">(上传新文件)</span>
                </label>
              </div>
            </div>
          </div>
          {error && <div className="text-red-500 dark:text-red-400 text-xs font-mono">{error}</div>}
          <div className="flex gap-2 pt-2">
            <button
              type="button"
              onClick={onCancel}
              className="flex-1 py-2 px-4 border border-zinc-200 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400 hover:text-zinc-900 dark:hover:text-zinc-100 hover:border-zinc-400 dark:hover:border-zinc-500 text-sm font-mono transition rounded"
            >
              取消
            </button>
            <button
              type="submit"
              disabled={loading}
              className="flex-1 py-2 px-4 bg-blue-600 hover:bg-blue-500 text-white text-sm font-mono disabled:opacity-50 transition rounded"
            >
              {loading ? "保存中..." : "保存"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

function SettingsModal({
  onClose,
  siteTitle,
  siteAnnouncement,
  isDark,
  onToggleTheme,
  isAdmin,
  onRefreshStorages,
  webdavEnabled,
  storages,
}: {
  onClose: () => void;
  siteTitle: string;
  siteAnnouncement: string;
  isDark: boolean;
  onToggleTheme: (e: React.MouseEvent) => void;
  isAdmin: boolean;
  onRefreshStorages: () => void;
  webdavEnabled: boolean;
  storages: StorageInfo[];
}) {
  const [activeTab, setActiveTab] = useState<'general' | 'webdav' | 'backup' | 'audit' | 'about'>('general');
  const [exporting, setExporting] = useState(false);
  const [importing, setImporting] = useState(false);
  const [importMode, setImportMode] = useState<'merge' | 'replace'>('merge');
  const [importResult, setImportResult] = useState<{ success: boolean; message: string } | null>(null);
  const [auditLogs, setAuditLogs] = useState<AuditLog[]>([]);
  const [auditLoading, setAuditLoading] = useState(false);
  const [auditError, setAuditError] = useState("");

  const handleExportBackup = async () => {
    setExporting(true);
    try {
      const res = await fetch("/api/storages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "export-backup" }),
      });

      if (res.ok) {
        const data = await res.json() as { backup: unknown };
        const blob = new Blob([JSON.stringify(data.backup, null, 2)], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `clist-backup-${new Date().toISOString().slice(0, 10)}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      } else {
        const data = await res.json() as { error?: string };
        alert(data.error || "导出失败");
      }
    } catch {
      alert("网络错误");
    } finally {
      setExporting(false);
    }
  };

  const handleImportBackup = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setImporting(true);
    setImportResult(null);

    try {
      const text = await file.text();
      const backup = JSON.parse(text);

      if (!backup.storages || !Array.isArray(backup.storages)) {
        setImportResult({ success: false, message: "无效的备份文件格式" });
        return;
      }

      const res = await fetch("/api/storages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "import-backup", backup, mode: importMode }),
      });

      const data = await res.json() as { success?: boolean; imported?: number; skipped?: number; errors?: string[]; error?: string };

      if (res.ok && data.success) {
        let message = `成功导入 ${data.imported} 个存储`;
        if (data.skipped && data.skipped > 0) {
          message += `，跳过 ${data.skipped} 个已存在的存储`;
        }
        if (data.errors && data.errors.length > 0) {
          message += `\n\n错误:\n${data.errors.join("\n")}`;
        }
        setImportResult({ success: true, message });
        onRefreshStorages();
      } else {
        setImportResult({ success: false, message: data.error || "导入失败" });
      }
    } catch (err) {
      setImportResult({ success: false, message: err instanceof Error ? err.message : "解析备份文件失败" });
    } finally {
      setImporting(false);
      e.target.value = "";
    }
  };

  const fetchAuditLogs = async () => {
    setAuditLoading(true);
    setAuditError("");
    try {
      const res = await fetch("/api/audit?limit=200");
      if (res.ok) {
        const data = await res.json() as { logs?: AuditLog[] };
        setAuditLogs(data.logs || []);
      } else {
        const data = await res.json() as { error?: string };
        setAuditError(data.error || "加载审计日志失败");
      }
    } catch {
      setAuditError("网络错误");
    } finally {
      setAuditLoading(false);
    }
  };

  useEffect(() => {
    if (activeTab === "audit" && isAdmin) {
      fetchAuditLogs();
    }
  }, [activeTab, isAdmin]);

  return (
    <div className="fixed inset-0 bg-black/50 dark:bg-black/70 backdrop-blur-sm flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700/50 w-full max-w-md rounded-xl shadow-2xl clist-modal-enter" onClick={e => e.stopPropagation()}>
        <div className="px-5 py-4 border-b border-zinc-200 dark:border-zinc-700/50 flex items-center justify-between">
          <span className="text-zinc-900 dark:text-zinc-100 font-medium text-sm">设置</span>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>
        </div>

        {/* Tabs */}
        <div className="flex border-b border-zinc-200 dark:border-zinc-700">
          <button
            onClick={() => setActiveTab('general')}
            className={`flex-1 px-4 py-2 text-xs font-mono transition ${
              activeTab === 'general'
                ? 'text-blue-500 border-b-2 border-blue-500'
                : 'text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300'
            }`}
          >
            常规
          </button>
          {isAdmin && (
            <button
              onClick={() => setActiveTab('webdav')}
              className={`flex-1 px-4 py-2 text-xs font-mono transition ${
                activeTab === 'webdav'
                  ? 'text-blue-500 border-b-2 border-blue-500'
                  : 'text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300'
              }`}
            >
              WebDAV
            </button>
          )}
          {isAdmin && (
            <button
              onClick={() => setActiveTab('backup')}
              className={`flex-1 px-4 py-2 text-xs font-mono transition ${
                activeTab === 'backup'
                  ? 'text-blue-500 border-b-2 border-blue-500'
                  : 'text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300'
              }`}
            >
              备份
            </button>
          )}
          {isAdmin && (
            <button
              onClick={() => setActiveTab('audit')}
              className={activeTab === 'audit' ? 'flex-1 px-4 py-2 text-xs font-mono transition text-blue-500 border-b-2 border-blue-500' : 'flex-1 px-4 py-2 text-xs font-mono transition text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300'}
            >
              审计
            </button>
          )}
          <button
            onClick={() => setActiveTab('about')}
            className={`flex-1 px-4 py-2 text-xs font-mono transition ${
              activeTab === 'about'
                ? 'text-blue-500 border-b-2 border-blue-500'
                : 'text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300'
            }`}
          >
            关于
          </button>
        </div>

        <div className="p-4">
          {activeTab === 'general' && (
            <div className="space-y-4">
              {/* Theme Setting */}
              <div className="flex items-center justify-between py-2">
                <div>
                  <div className="text-sm text-zinc-900 dark:text-zinc-100 font-mono">主题模式</div>
                  <div className="text-xs text-zinc-500">切换亮色或暗色主题</div>
                </div>
                <button
                  onClick={onToggleTheme}
                  className="px-3 py-1.5 text-xs font-mono rounded border border-zinc-200 dark:border-zinc-700 bg-zinc-50 dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-700 transition"
                >
                  {isDark ? '☀ 亮色' : '☾ 暗色'}
                </button>
              </div>

              {/* Announcement */}
              {siteAnnouncement && (
                <div className="border-t border-zinc-200 dark:border-zinc-700 pt-4">
                  <div className="text-sm text-zinc-900 dark:text-zinc-100 font-mono mb-2 flex items-center gap-2">
                    <span className="text-yellow-500">📢</span> 公告
                  </div>
                  <div className="text-xs text-zinc-600 dark:text-zinc-400 font-mono whitespace-pre-wrap bg-zinc-50 dark:bg-zinc-800 p-3 rounded border border-zinc-200 dark:border-zinc-700 max-h-32 overflow-y-auto">
                    {siteAnnouncement}
                  </div>
                </div>
              )}
            </div>
          )}

          {activeTab === 'webdav' && isAdmin && (
            <div className="space-y-4">
              {/* WebDAV Status */}
              <div className="flex items-center justify-between py-2">
                <div>
                  <div className="text-sm text-zinc-900 dark:text-zinc-100 font-mono">WebDAV 服务</div>
                  <div className="text-xs text-zinc-500">通过 WebDAV 协议访问存储</div>
                </div>
                <span className={`px-2 py-1 text-xs font-mono rounded ${
                  webdavEnabled 
                    ? 'bg-green-100 dark:bg-green-900/30 text-green-700 dark:text-green-400' 
                    : 'bg-zinc-100 dark:bg-zinc-800 text-zinc-500'
                }`}>
                  {webdavEnabled ? '已启用' : '未启用'}
                </span>
              </div>

              {webdavEnabled ? (
                <>
                  {/* WebDAV URL */}
                  <div className="border-t border-zinc-200 dark:border-zinc-700 pt-4">
                    <div className="text-sm text-zinc-900 dark:text-zinc-100 font-mono mb-2">访问地址</div>
                    <div className="text-xs text-zinc-500 mb-3">
                      使用 WebDAV 客户端连接以下地址访问存储
                    </div>
                    <div className="bg-zinc-50 dark:bg-zinc-800 p-3 rounded border border-zinc-200 dark:border-zinc-700">
                      <div className="text-xs text-zinc-500 mb-1 font-mono">根目录 (所有存储):</div>
                      <code className="text-sm text-blue-600 dark:text-blue-400 font-mono break-all">
                        {typeof window !== 'undefined' ? `${window.location.origin}/dav/0/` : '/dav/0/'}
                      </code>
                    </div>
                  </div>

                  {/* Storage List with WebDAV URLs */}
                  {storages.length > 0 && (
                    <div className="border-t border-zinc-200 dark:border-zinc-700 pt-4">
                      <div className="text-sm text-zinc-900 dark:text-zinc-100 font-mono mb-2">存储访问地址</div>
                      <div className="space-y-2 max-h-48 overflow-y-auto">
                        {storages.map((storage) => (
                          <div key={storage.id} className="bg-zinc-50 dark:bg-zinc-800 p-2 rounded border border-zinc-200 dark:border-zinc-700">
                            <div className="text-xs text-zinc-700 dark:text-zinc-300 font-mono mb-1">{storage.name}</div>
                            <code className="text-xs text-blue-600 dark:text-blue-400 font-mono break-all">
                              {typeof window !== 'undefined' ? `${window.location.origin}/dav/${storage.id}/` : `/dav/${storage.id}/`}
                            </code>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Authentication Info */}
                  <div className="border-t border-zinc-200 dark:border-zinc-700 pt-4">
                    <div className="text-sm text-zinc-900 dark:text-zinc-100 font-mono mb-2">认证方式</div>
                    <div className="text-xs text-zinc-600 dark:text-zinc-400 font-mono space-y-1">
                      <p>• 协议: HTTP Basic Authentication</p>
                      <p>• 用户名/密码: 使用 WEBDAV_USERNAME/WEBDAV_PASSWORD 环境变量配置</p>
                      <p>• 默认: 使用管理员账号密码 (ADMIN_USERNAME/ADMIN_PASSWORD)</p>
                    </div>
                  </div>

                  {/* Usage Tips */}
                  <div className="border-t border-zinc-200 dark:border-zinc-700 pt-4">
                    <div className="text-sm text-zinc-900 dark:text-zinc-100 font-mono mb-2 flex items-center gap-2">
                      <span className="text-blue-500">💡</span> 使用提示
                    </div>
                    <div className="text-xs text-zinc-600 dark:text-zinc-400 font-mono space-y-1">
                      <p>• Windows: 映射网络驱动器，输入 WebDAV 地址</p>
                      <p>• macOS: Finder → 前往 → 连接服务器</p>
                      <p>• Linux: 使用 davfs2 或文件管理器</p>
                      <p>• 移动端: 使用支持 WebDAV 的文件管理 App</p>
                    </div>
                  </div>
                </>
              ) : (
                <div className="border-t border-zinc-200 dark:border-zinc-700 pt-4">
                  <div className="text-xs text-zinc-500 font-mono space-y-2">
                    <p>WebDAV 服务未启用。要启用 WebDAV，请在 Cloudflare Workers 环境变量中设置:</p>
                    <div className="bg-zinc-50 dark:bg-zinc-800 p-3 rounded border border-zinc-200 dark:border-zinc-700 mt-2">
                      <code className="text-xs text-zinc-700 dark:text-zinc-300">WEBDAV_ENABLED = "true"</code>
                    </div>
                    <p className="mt-2">可选配置:</p>
                    <div className="bg-zinc-50 dark:bg-zinc-800 p-3 rounded border border-zinc-200 dark:border-zinc-700">
                      <code className="text-xs text-zinc-700 dark:text-zinc-300 block">WEBDAV_USERNAME = "your_username"</code>
                      <code className="text-xs text-zinc-700 dark:text-zinc-300 block">WEBDAV_PASSWORD = "your_password"</code>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {activeTab === 'backup' && isAdmin && (
            <div className="space-y-4">
              {/* Export Section */}
              <div>
                <div className="text-sm text-zinc-900 dark:text-zinc-100 font-mono mb-2">导出备份</div>
                <div className="text-xs text-zinc-500 mb-3">
                  导出所有存储配置到 JSON 文件，包含连接凭证信息。
                </div>
                <button
                  onClick={handleExportBackup}
                  disabled={exporting}
                  className="w-full py-2 px-4 bg-blue-600 hover:bg-blue-500 text-white text-sm font-mono disabled:opacity-50 transition rounded"
                >
                  {exporting ? "导出中..." : "导出备份文件"}
                </button>
              </div>

              {/* Import Section */}
              <div className="border-t border-zinc-200 dark:border-zinc-700 pt-4">
                <div className="text-sm text-zinc-900 dark:text-zinc-100 font-mono mb-2">恢复备份</div>
                <div className="text-xs text-zinc-500 mb-3">
                  从备份文件恢复存储配置。
                </div>

                {/* Import Mode Selection */}
                <div className="mb-3">
                  <div className="text-xs text-zinc-500 mb-2 font-mono">导入模式:</div>
                  <div className="flex gap-4">
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="radio"
                        name="importMode"
                        value="merge"
                        checked={importMode === 'merge'}
                        onChange={() => setImportMode('merge')}
                        className="w-4 h-4"
                      />
                      <span className="text-sm text-zinc-700 dark:text-zinc-300 font-mono">合并</span>
                      <span className="text-xs text-zinc-500">(保留现有)</span>
                    </label>
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="radio"
                        name="importMode"
                        value="replace"
                        checked={importMode === 'replace'}
                        onChange={() => setImportMode('replace')}
                        className="w-4 h-4"
                      />
                      <span className="text-sm text-zinc-700 dark:text-zinc-300 font-mono">替换</span>
                      <span className="text-xs text-zinc-500">(清空现有)</span>
                    </label>
                  </div>
                </div>

                <label className={`block w-full py-2 px-4 text-center border-2 border-dashed border-zinc-300 dark:border-zinc-600 hover:border-blue-500 dark:hover:border-blue-500 text-sm font-mono cursor-pointer transition rounded ${importing ? 'opacity-50 pointer-events-none' : ''}`}>
                  {importing ? "导入中..." : "选择备份文件"}
                  <input
                    type="file"
                    accept=".json"
                    onChange={handleImportBackup}
                    className="hidden"
                    disabled={importing}
                  />
                </label>

                {/* Import Result */}
                {importResult && (
                  <div className={`mt-3 p-3 rounded text-xs font-mono whitespace-pre-wrap ${
                    importResult.success
                      ? 'bg-green-50 dark:bg-green-900/20 text-green-700 dark:text-green-400 border border-green-200 dark:border-green-800'
                      : 'bg-red-50 dark:bg-red-900/20 text-red-700 dark:text-red-400 border border-red-200 dark:border-red-800'
                  }`}>
                    {importResult.message}
                  </div>
                )}
              </div>

              {/* Warning */}
              <div className="border-t border-zinc-200 dark:border-zinc-700 pt-4">
                <div className="text-xs text-yellow-600 dark:text-yellow-500 font-mono flex items-start gap-2">
                  <span>⚠</span>
                  <span>备份文件包含敏感凭证信息，请妥善保管。</span>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'audit' && isAdmin && (
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <div className="text-sm text-zinc-900 dark:text-zinc-100 font-mono">审计日志</div>
                <button
                  onClick={fetchAuditLogs}
                  disabled={auditLoading}
                  className="px-3 py-1 text-xs font-mono rounded border border-zinc-200 dark:border-zinc-700 bg-zinc-50 dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-700 disabled:opacity-50 transition"
                >
                  {auditLoading ? '加载中...' : '刷新'}
                </button>
              </div>
              {auditError && (
                <div className="text-xs text-red-500 dark:text-red-400 font-mono">{auditError}</div>
              )}
              {!auditError && auditLogs.length === 0 && !auditLoading && (
                <div className="text-xs text-zinc-500 font-mono">暂无日志</div>
              )}
              {auditLogs.length > 0 && (
                <div className="space-y-2 max-h-72 overflow-y-auto">
                  {auditLogs.map((log) => (
                    <div key={log.id} className="border border-zinc-200 dark:border-zinc-700 rounded p-2 bg-zinc-50 dark:bg-zinc-800/50">
                      <div className="flex items-center justify-between">
                        <span className="text-xs text-zinc-500 font-mono">{formatDate(log.createdAt)}</span>
                        <span className="text-[11px] text-zinc-400 font-mono">{log.userType}</span>
                      </div>
                      <div className="text-xs text-zinc-800 dark:text-zinc-200 font-mono">{log.action}</div>
                      <div className="text-[11px] text-zinc-500 font-mono">
                        {log.storageId ? `storage #${log.storageId}` : 'storage -'}
                        {log.path ? ` / ${log.path}` : ''}
                      </div>
                      {log.detail && (
                        <div className="text-[11px] text-zinc-500 dark:text-zinc-400 font-mono mt-1 break-all">{log.detail}</div>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {activeTab === 'about' && (
            <div className="space-y-4">
              <div className="text-center py-4">
                <div className="text-3xl font-bold text-zinc-900 dark:text-zinc-100 font-mono mb-1">{siteTitle}</div>
                <div className="text-xs text-zinc-500 font-mono">v1.2.0</div>
              </div>
              <div className="text-xs text-zinc-600 dark:text-zinc-400 font-mono space-y-2">
                <p>S3 兼容存储聚合服务</p>
                <p className="text-zinc-500">支持: AWS S3 / Cloudflare R2 / 阿里云 OSS / 腾讯云 COS / MinIO / WebDAV / OneDrive / Google Drive / 阿里云盘 / 百度网盘</p>
                <p>作者: ooyyh</p>
                <p>联系方式: 3266940347@qq.com</p>
              </div>
              <div className="border-t border-zinc-200 dark:border-zinc-800 pt-4 text-xs text-zinc-500 font-mono">
                <p>Powered by Cloudflare Workers && ooyyh</p>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function AnnouncementModal({ announcement, onClose }: { announcement: string; onClose: () => void }) {
  return (
    <div className="fixed inset-0 bg-black/50 dark:bg-black/70 backdrop-blur-sm flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700/50 w-full max-w-lg rounded-xl shadow-2xl clist-modal-enter" onClick={e => e.stopPropagation()}>
        <div className="px-5 py-4 border-b border-zinc-200 dark:border-zinc-700/50 flex items-center justify-between">
          <span className="text-zinc-900 dark:text-zinc-100 font-medium text-sm flex items-center gap-2">
            <span className="text-yellow-500">📢</span> 公告
          </span>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>
        </div>
        <div className="p-5">
          <p className="text-sm text-zinc-700 dark:text-zinc-300 whitespace-pre-wrap leading-relaxed">
            {announcement}
          </p>
        </div>
        <div className="px-5 pb-5">
          <button
            onClick={onClose}
            className="w-full py-2.5 px-4 bg-blue-600 hover:bg-blue-500 text-white text-sm font-medium transition rounded-lg shadow-sm"
          >
            我知道了
          </button>
        </div>
      </div>
    </div>
  );
}

interface StorageStats {
  totalSize: number;
  fileCount: number;
  folderCount: number;
  typeDistribution: Record<string, { count: number; size: number }>;
}

const chartColors = ["#2563eb", "#10b981", "#f59e0b", "#8b5cf6", "#ef4444", "#06b6d4", "#84cc16", "#ec4899", "#64748b", "#14b8a6"];

function buildConicGradient(items: Array<{ percentage: number; color: string }>): string {
  if (items.length === 0) {
    return "conic-gradient(#d4d4d8 0deg 360deg)";
  }
  let start = 0;
  const stops = items.map((item) => {
    const end = start + item.percentage * 3.6;
    const stop = `${item.color} ${start.toFixed(2)}deg ${end.toFixed(2)}deg`;
    start = end;
    return stop;
  });
  return `conic-gradient(${stops.join(", ")})`;
}

function StorageStatsModal({ storage, onClose }: { storage: StorageInfo; onClose: () => void }) {
  const [stats, setStats] = useState<StorageStats | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    const fetchStats = async () => {
      setLoading(true);
      setError("");
      try {
        const res = await fetch(`/api/storage-stats/${storage.id}`);
        if (res.ok) {
          const data = (await res.json()) as { stats: StorageStats };
          setStats(data.stats);
        } else {
          const data = (await res.json()) as { error?: string };
          setError(data.error || "获取统计信息失败");
        }
      } catch {
        setError("网络错误");
      } finally {
        setLoading(false);
      }
    };
    fetchStats();
  }, [storage.id]);

  const sortedTypes = stats
    ? Object.entries(stats.typeDistribution)
        .sort((a, b) => b[1].size - a[1].size)
    : [];
  const chartItems = stats
    ? (() => {
        const topTypes = sortedTypes.slice(0, 10);
        const items = topTypes.map(([ext, data], index) => ({
          ext,
          count: data.count,
          size: data.size,
          percentage: stats.totalSize > 0 ? (data.size / stats.totalSize) * 100 : 0,
          color: chartColors[index % chartColors.length],
        }));
        const shownSize = topTypes.reduce((sum, [, data]) => sum + data.size, 0);
        const shownCount = topTypes.reduce((sum, [, data]) => sum + data.count, 0);
        const restSize = stats.totalSize - shownSize;
        const restCount = stats.fileCount - shownCount;
        if (restSize > 0 || restCount > 0) {
          items.push({
            ext: "other",
            count: Math.max(0, restCount),
            size: Math.max(0, restSize),
            percentage: stats.totalSize > 0 ? (Math.max(0, restSize) / stats.totalSize) * 100 : 0,
            color: chartColors[items.length % chartColors.length],
          });
        }
        return items;
      })()
    : [];
  const donutGradient = buildConicGradient(chartItems.map(({ percentage, color }) => ({ percentage, color })));
  const dominantType = chartItems[0];

  return (
    <div className="fixed inset-0 bg-black/50 dark:bg-black/70 backdrop-blur-sm flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700/50 w-full max-w-2xl lg:max-w-3xl max-h-[80vh] lg:max-h-[84vh] rounded-xl shadow-2xl flex flex-col clist-modal-enter" onClick={e => e.stopPropagation()}>
        <div className="px-5 py-4 border-b border-zinc-200 dark:border-zinc-700/50 flex items-center justify-between shrink-0">
          <span className="text-zinc-900 dark:text-zinc-100 font-medium text-sm flex items-center gap-2">
            <span className="lg:hidden text-blue-500">📊</span>
            <span className="hidden lg:inline-flex h-7 w-7 items-center justify-center rounded-md border border-blue-200 bg-blue-50 text-blue-600 shadow-sm dark:border-blue-400/30 dark:bg-blue-500/10 dark:text-blue-300">
              <StatsIcon className="h-[18px] w-[18px]" />
            </span>
            存储统计 - {storage.name}
          </span>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>
        </div>
        <div className="flex-1 overflow-y-auto p-4">
          {loading ? (
            <div className="flex items-center justify-center py-8">
              <span className="text-zinc-400 dark:text-zinc-500 font-mono text-sm">正在统计中，请稍候...</span>
            </div>
          ) : error ? (
            <div className="flex items-center justify-center py-8">
              <span className="text-red-500 font-mono text-sm">{error}</span>
            </div>
          ) : stats ? (
            <div className="space-y-4">
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div className="bg-zinc-50 dark:bg-zinc-800 p-4 rounded border border-zinc-200 dark:border-zinc-700">
                  <div className="text-xs text-zinc-500 font-mono mb-1">总大小</div>
                  <div className="text-2xl font-mono text-zinc-900 dark:text-zinc-100">{formatBytes(stats.totalSize)}</div>
                </div>
                <div className="bg-zinc-50 dark:bg-zinc-800 p-4 rounded border border-zinc-200 dark:border-zinc-700">
                  <div className="text-xs text-zinc-500 font-mono mb-1">文件数量</div>
                  <div className="text-2xl font-mono text-zinc-900 dark:text-zinc-100">{stats.fileCount.toLocaleString()}</div>
                </div>
                <div className="bg-zinc-50 dark:bg-zinc-800 p-4 rounded border border-zinc-200 dark:border-zinc-700">
                  <div className="text-xs text-zinc-500 font-mono mb-1">文件夹数量</div>
                  <div className="text-2xl font-mono text-zinc-900 dark:text-zinc-100">{stats.folderCount.toLocaleString()}</div>
                </div>
              </div>

              {sortedTypes.length > 0 && (
                <>
                  <div className="grid grid-cols-1 lg:grid-cols-[260px_1fr] gap-3">
                    <div className="bg-zinc-50 dark:bg-zinc-800 p-4 rounded border border-zinc-200 dark:border-zinc-700">
                      <div className="flex items-center justify-between mb-3">
                        <div className="text-xs text-zinc-500 font-mono">容量构成</div>
                        <div className="text-[11px] text-zinc-400 dark:text-zinc-500 font-mono">Top {chartItems.length}</div>
                      </div>
                      <div className="flex items-center justify-center">
                        <div
                          className="relative h-40 w-40 rounded-full shadow-inner"
                          style={{ background: donutGradient }}
                          aria-label="文件类型容量环形图"
                        >
                          <div className="absolute inset-5 rounded-full bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 flex flex-col items-center justify-center">
                            <div className="text-[11px] text-zinc-500 font-mono">主类型</div>
                            <div className="text-xl text-zinc-900 dark:text-zinc-100 font-mono">{dominantType ? `.${dominantType.ext}` : "-"}</div>
                            <div className="text-xs text-zinc-500 font-mono">{dominantType ? `${dominantType.percentage.toFixed(1)}%` : "0%"}</div>
                          </div>
                        </div>
                      </div>
                    </div>

                    <div className="bg-zinc-50 dark:bg-zinc-800 p-4 rounded border border-zinc-200 dark:border-zinc-700">
                      <div className="text-xs text-zinc-500 font-mono mb-3">类型占比</div>
                      <div className="h-4 w-full overflow-hidden rounded-full bg-zinc-200 dark:bg-zinc-700 flex">
                        {chartItems.map((item) => (
                          <div
                            key={item.ext}
                            title={`.${item.ext} ${item.percentage.toFixed(1)}%`}
                            style={{ width: `${Math.max(item.percentage, 1)}%`, backgroundColor: item.color }}
                          />
                        ))}
                      </div>
                      <div className="mt-4 grid grid-cols-2 sm:grid-cols-3 gap-2">
                        {chartItems.slice(0, 6).map((item) => (
                          <div key={item.ext} className="min-w-0 rounded border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-900 px-2 py-2">
                            <div className="flex items-center gap-2 min-w-0">
                              <span className="h-2.5 w-2.5 rounded-full shrink-0" style={{ backgroundColor: item.color }} />
                              <span className="truncate text-xs text-zinc-700 dark:text-zinc-300 font-mono">.{item.ext}</span>
                            </div>
                            <div className="mt-1 text-[11px] text-zinc-500 font-mono">{formatBytes(item.size)} · {item.percentage.toFixed(1)}%</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>

                  <div className="bg-zinc-50 dark:bg-zinc-800 p-4 rounded border border-zinc-200 dark:border-zinc-700">
                    <div className="text-sm text-zinc-900 dark:text-zinc-100 font-mono mb-3">文件类型排行</div>
                    <div className="space-y-2.5">
                      {chartItems.map((item) => (
                        <div key={item.ext} className="grid grid-cols-[minmax(48px,72px)_minmax(0,1fr)_minmax(84px,112px)] items-center gap-2 sm:gap-3 text-xs font-mono">
                          <div className="truncate text-zinc-700 dark:text-zinc-300">.{item.ext}</div>
                          <div className="h-3 rounded-full bg-zinc-200 dark:bg-zinc-700 overflow-hidden">
                            <div
                              className="h-full rounded-full"
                              style={{ width: `${Math.max(item.percentage, 1)}%`, backgroundColor: item.color }}
                            />
                          </div>
                          <div className="text-right text-zinc-500">
                            {formatBytes(item.size)} · {item.count.toLocaleString()}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </>
              )}

              {stats.fileCount === 0 && (
                <div className="text-center py-8">
                  <span className="text-zinc-400 dark:text-zinc-500 font-mono text-sm">此存储为空</span>
                </div>
              )}
            </div>
          ) : null}
        </div>
        <div className="px-4 py-3 border-t border-zinc-200 dark:border-zinc-700 shrink-0">
          <button
            onClick={onClose}
            className="w-full py-2 px-4 bg-blue-600 hover:bg-blue-500 text-white text-sm font-mono transition rounded"
          >
            关闭
          </button>
        </div>
      </div>
    </div>
  );
}

interface ReleaseItem {
  version: string;
  name: string;
  body: string;
  publishedAt: string;
  url: string;
  isPrerelease: boolean;
  author: string;
}

function ChangelogModal({ onClose }: { onClose: () => void }) {
  const [releases, setReleases] = useState<ReleaseItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    const fetchReleases = async () => {
      try {
        const res = await fetch("/api/changelog");
        if (res.ok) {
          const data = await res.json() as { releases: ReleaseItem[] };
          setReleases(data.releases);
        } else {
          setError("获取更新日志失败");
        }
      } catch {
        setError("网络错误");
      } finally {
        setLoading(false);
      }
    };
    fetchReleases();
  }, []);

  const formatDate = (dateStr: string) => {
    const date = new Date(dateStr);
    return date.toLocaleDateString("zh-CN", { year: "numeric", month: "long", day: "numeric" });
  };

  const parseBody = (body: string) => {
    // Parse the changelog body and highlight different types
    return body.split("\n").map((line, i) => {
      const trimmed = line.trim();
      if (!trimmed) return null;

      let colorClass = "text-zinc-600 dark:text-zinc-400";
      if (trimmed.toLowerCase().startsWith("#update") || trimmed.toLowerCase().startsWith("update")) {
        colorClass = "text-blue-600 dark:text-blue-400";
      } else if (trimmed.toLowerCase().startsWith("#fix") || trimmed.toLowerCase().startsWith("fix")) {
        colorClass = "text-green-600 dark:text-green-400";
      } else if (trimmed.toLowerCase().startsWith("#breaking") || trimmed.toLowerCase().startsWith("breaking")) {
        colorClass = "text-red-600 dark:text-red-400";
      } else if (trimmed.toLowerCase().startsWith("#new") || trimmed.toLowerCase().startsWith("new")) {
        colorClass = "text-purple-600 dark:text-purple-400";
      }

      return (
        <div key={i} className={`${colorClass} text-sm font-mono`}>
          {trimmed}
        </div>
      );
    });
  };

  return (
    <div className="fixed inset-0 bg-black/50 dark:bg-black/70 backdrop-blur-sm flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700/50 w-full max-w-2xl max-h-[80vh] rounded-xl shadow-2xl flex flex-col clist-modal-enter" onClick={e => e.stopPropagation()}>
        <div className="px-5 py-4 border-b border-zinc-200 dark:border-zinc-700/50 flex items-center justify-between shrink-0">
          <span className="text-zinc-900 dark:text-zinc-100 font-medium text-sm flex items-center gap-2">
            <span className="text-blue-500">📋</span> 更新日志
          </span>
          <button onClick={onClose} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>
        </div>
        <div className="flex-1 overflow-y-auto p-4">
          {loading ? (
            <div className="flex items-center justify-center py-8">
              <span className="text-zinc-400 dark:text-zinc-500 font-mono text-sm">加载中...</span>
            </div>
          ) : error ? (
            <div className="flex items-center justify-center py-8">
              <span className="text-red-500 font-mono text-sm">{error}</span>
            </div>
          ) : releases.length === 0 ? (
            <div className="flex items-center justify-center py-8">
              <span className="text-zinc-400 dark:text-zinc-500 font-mono text-sm">暂无更新日志</span>
            </div>
          ) : (
            <div className="space-y-6">
              {releases.map((release, idx) => (
                <div key={release.version} className="relative">
                  {idx > 0 && <div className="absolute -top-3 left-0 right-0 border-t border-zinc-200 dark:border-zinc-700" />}
                  <div className="flex items-center gap-3 mb-2">
                    <span className={`px-2 py-0.5 text-xs font-mono rounded ${
                      idx === 0
                        ? "bg-blue-100 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400"
                        : "bg-zinc-100 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-400"
                    }`}>
                      {release.version}
                    </span>
                    {idx === 0 && (
                      <span className="px-2 py-0.5 text-xs font-mono rounded bg-green-100 dark:bg-green-900/30 text-green-600 dark:text-green-400">
                        Latest
                      </span>
                    )}
                    {release.isPrerelease && (
                      <span className="px-2 py-0.5 text-xs font-mono rounded bg-yellow-100 dark:bg-yellow-900/30 text-yellow-600 dark:text-yellow-400">
                        Pre-release
                      </span>
                    )}
                    <span className="text-xs text-zinc-400 dark:text-zinc-500 font-mono">
                      {formatDate(release.publishedAt)}
                    </span>
                  </div>
                  {release.name && release.name !== release.version && (
                    <h3 className="text-sm font-mono text-zinc-800 dark:text-zinc-200 mb-2">{release.name}</h3>
                  )}
                  <div className="space-y-1 pl-2 border-l-2 border-zinc-200 dark:border-zinc-700">
                    {parseBody(release.body)}
                  </div>
                  <a
                    href={release.url}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-block mt-2 text-xs text-blue-500 hover:text-blue-400 font-mono"
                  >
                    查看详情 →
                  </a>
                </div>
              ))}
            </div>
          )}
        </div>
        <div className="px-4 py-3 border-t border-zinc-200 dark:border-zinc-700 shrink-0">
          <button
            onClick={onClose}
            className="w-full py-2 px-4 bg-blue-600 hover:bg-blue-500 text-white text-sm font-mono transition rounded"
          >
            关闭
          </button>
        </div>
      </div>
    </div>
  );
}

function FileBrowser({ storage, isAdmin, isDark, chunkSizeMB }: { storage: StorageInfo; isAdmin: boolean; isDark: boolean; chunkSizeMB: number }) {
  // Permission checks
  const canList = isAdmin || storage.guestList;
  const canDownload = isAdmin || storage.guestDownload;
  const canUpload = isAdmin || storage.guestUpload;

  const [path, setPath] = useState("");
  const [objects, setObjects] = useState<S3Object[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [uploadProgress, setUploadProgress] = useState<{
    name: string;
    progress: number;
    currentPart?: number;
    totalParts?: number;
    speed?: number; // bytes per second
    loaded?: number;
    total?: number;
  } | null>(null);
  const [previewFile, setPreviewFile] = useState<S3Object | null>(null);
  const [showNewFolderInput, setShowNewFolderInput] = useState(false);
  const [newFolderName, setNewFolderName] = useState("");
  const [creatingFolder, setCreatingFolder] = useState(false);
  const [showOfflineDownload, setShowOfflineDownload] = useState(false);
  const [offlineUrl, setOfflineUrl] = useState("");
  const [offlineFilename, setOfflineFilename] = useState("");
  const [offlineDownloading, setOfflineDownloading] = useState(false);
  const [selectedKeys, setSelectedKeys] = useState<Set<string>>(new Set());
  const [deleting, setDeleting] = useState(false);
  const [renameTarget, setRenameTarget] = useState<S3Object | null>(null);
  const [renameValue, setRenameValue] = useState("");
  const [renaming, setRenaming] = useState(false);
  const [moveTarget, setMoveTarget] = useState<S3Object | null>(null);
  const [moveDestPath, setMoveDestPath] = useState("");
  const [moving, setMoving] = useState(false);
  const [allFolders, setAllFolders] = useState<string[]>([]);
  const [shareTarget, setShareTarget] = useState<S3Object | null>(null);
  const [shareToken, setShareToken] = useState("");
  const [shareUrl, setShareUrl] = useState("");
  const [shareExpireHours, setShareExpireHours] = useState(0);
  const [creatingShare, setCreatingShare] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [searchResults, setSearchResults] = useState<S3Object[] | null>(null);
  const [searchLoading, setSearchLoading] = useState(false);
  // Batch operation states
  const [batchMode, setBatchMode] = useState<"move" | "copy" | null>(null);
  const [batchDestPath, setBatchDestPath] = useState("");
  const [batchProcessing, setBatchProcessing] = useState(false);
  const [batchProgress, setBatchProgress] = useState<{ current: number; total: number } | null>(null);

  useEffect(() => {
    setPath("");
    setSearchQuery("");
  }, [storage.id]);

  useEffect(() => {
    loadFiles();
    setSelectedKeys(new Set()); // Clear selection on path change
  }, [storage.id, path]);

  const loadFiles = async () => {
    setLoading(true);
    setError("");

    try {
      const res = await fetch(`/api/files/${storage.id}/${path}?action=list`);
      if (res.ok) {
        const data = (await res.json()) as { objects?: S3Object[] };
        setObjects(data.objects || []);
      } else {
        const data = (await res.json()) as { error?: string };
        setError(data.error || "加载失败");
      }
    } catch {
      setError("网络错误");
    } finally {
      setLoading(false);
    }
  };

  const navigateTo = (newPath: string) => {
    setPath(newPath.replace(/^\//, "").replace(/\/$/, ""));
  };

  const goUp = () => {
    const parts = path.split("/").filter(Boolean);
    parts.pop();
    setPath(parts.join("/"));
  };

  const downloadFile = (key: string) => {
    window.open(`/api/files/${storage.id}/${key}?action=download`, "_blank");
  };

  const deleteFile = async (key: string) => {
    if (!confirm(`确定删除 ${key}?`)) return;
    try {
      const res = await fetch(`/api/files/${storage.id}/${key}`, { method: "DELETE" });
      if (res.ok) {
        loadFiles();
      } else {
        const data = (await res.json()) as { error?: string };
        alert(data.error || "删除失败");
      }
    } catch {
      alert("网络错误");
    }
  };

  const deleteFolder = async (key: string, name: string) => {
    if (!confirm(`确定删除文件夹 "${name}" 及其所有内容?`)) return;
    try {
      const res = await fetch(`/api/files/${storage.id}/${key}?action=rmdir`, { method: "DELETE" });
      if (res.ok) {
        loadFiles();
      } else {
        const data = (await res.json()) as { error?: string };
        alert(data.error || "删除失败");
      }
    } catch {
      alert("网络错误");
    }
  };

  const startRename = (obj: S3Object) => {
    setRenameTarget(obj);
    setRenameValue(obj.name);
  };

  const handleRename = async () => {
    if (!renameTarget || !renameValue.trim()) return;
    if (renameValue.includes("/")) {
      alert("名称不能包含 /");
      return;
    }
    if (renameValue === renameTarget.name) {
      setRenameTarget(null);
      return;
    }

    setRenaming(true);
    try {
      const key = renameTarget.isDirectory ? renameTarget.key : renameTarget.key;
      const res = await fetch(`/api/files/${storage.id}/${key}?action=rename`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ newName: renameValue.trim() }),
      });
      if (res.ok) {
        setRenameTarget(null);
        loadFiles();
      } else {
        const data = (await res.json()) as { error?: string };
        alert(data.error || "重命名失败");
      }
    } catch {
      alert("网络错误");
    } finally {
      setRenaming(false);
    }
  };

  const loadAllFolders = async () => {
    const folders: string[] = [""];
    const MAX_DEPTH = 5;
    const listRecursive = async (prefix: string, depth: number) => {
      if (depth >= MAX_DEPTH) return;
      try {
        const res = await fetch(`/api/files/${storage.id}/${prefix}?action=list`);
        if (res.ok) {
          const data = (await res.json()) as { objects?: S3Object[] };
          for (const obj of data.objects || []) {
            if (obj.isDirectory) {
              folders.push(obj.key);
              await listRecursive(obj.key, depth + 1);
            }
          }
        }
      } catch {
        // Ignore errors
      }
    };
    await listRecursive("", 0);
    setAllFolders(folders);
  };

  const startMove = async (obj: S3Object) => {
    setMoveTarget(obj);
    setMoveDestPath("");
    await loadAllFolders();
  };

  const handleMove = async () => {
    if (!moveTarget) return;

    setMoving(true);
    try {
      const key = moveTarget.isDirectory ? moveTarget.key : moveTarget.key;
      const res = await fetch(`/api/files/${storage.id}/${key}?action=move`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ destPath: moveDestPath }),
      });
      if (res.ok) {
        setMoveTarget(null);
        loadFiles();
      } else {
        const data = (await res.json()) as { error?: string };
        alert(data.error || "移动失败");
      }
    } catch {
      alert("网络错误");
    } finally {
      setMoving(false);
    }
  };

  const startShare = (obj: S3Object) => {
    setShareTarget(obj);
    setShareToken("");
    setShareUrl("");
    setShareExpireHours(0);
  };

  const handleCreateShare = async () => {
    if (!shareTarget) return;

    setCreatingShare(true);
    try {
      let expiresAt: string | undefined;
      if (shareExpireHours > 0) {
        const expireDate = new Date();
        expireDate.setHours(expireDate.getHours() + shareExpireHours);
        expiresAt = expireDate.toISOString();
      }

      const res = await fetch("/api/shares", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          storageId: storage.id,
          filePath: shareTarget.key,
          isDirectory: shareTarget.isDirectory,
          expiresAt,
        }),
      });

      if (res.ok) {
        const data = (await res.json()) as { share: { shareToken: string }; shareUrl: string };
        setShareToken(data.share.shareToken);
        setShareUrl(data.shareUrl);
      } else {
        const data = (await res.json()) as { error?: string };
        alert(data.error || "创建分享链接失败");
      }
    } catch {
      alert("网络错误");
    } finally {
      setCreatingShare(false);
    }
  };

  const copyToClipboard = (text: string) => {
    // 尝试现代 Clipboard API（需要 HTTPS）
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(text).then(() => {
        alert("已复制到剪贴板");
      }).catch(() => {
        fallbackCopy(text);
      });
    } else {
      fallbackCopy(text);
    }
  };

  const fallbackCopy = (text: string) => {
    // 兼容方案：使用临时 textarea + execCommand
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.left = "-9999px";
    textarea.style.top = "-9999px";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    try {
      document.execCommand("copy");
      alert("已复制到剪贴板");
    } catch {
      alert("复制失败，请手动复制");
    } finally {
      document.body.removeChild(textarea);
    }
  };

  const toggleSelect = (key: string) => {
    setSelectedKeys((prev) => {
      const next = new Set(prev);
      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }
      return next;
    });
  };

  const toggleSelectAll = () => {
    setSelectedKeys((prev) => {
      const next = new Set(prev);
      if (allVisibleSelected) {
        for (const obj of visibleObjects) {
          next.delete(obj.key);
        }
      } else {
        for (const obj of visibleObjects) {
          next.add(obj.key);
        }
      }
      return next;
    });
  };

  const handleBatchDelete = async () => {
    if (selectedKeys.size === 0) return;

    const selected = objects.filter((obj) => selectedKeys.has(obj.key));
    const folders = selected.filter((obj) => obj.isDirectory);
    const files = selected.filter((obj) => !obj.isDirectory);

    const msg = folders.length > 0
      ? `确定删除 ${files.length} 个文件和 ${folders.length} 个文件夹（含其中所有内容）?`
      : `确定删除 ${files.length} 个文件?`;

    if (!confirm(msg)) return;

    setDeleting(true);
    try {
      const res = await fetch(`/api/files/${storage.id}/${path}?action=batch`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          operation: "delete",
          items: Array.from(selectedKeys),
        }),
      });
      const data = (await res.json()) as { success?: string[]; failed?: { key: string; error: string }[] };
      if (res.ok) {
        const failedCount = data.failed?.length || 0;
        if (failedCount > 0) {
          alert(`删除完成，${failedCount} 个项目删除失败`);
        }
      } else {
        alert((data as { error?: string }).error || "批量删除失败");
      }
      setSelectedKeys(new Set());
      loadFiles();
    } catch {
      alert("网络错误");
    } finally {
      setDeleting(false);
    }
  };

  const handleBatchMoveCopy = async () => {
    if (selectedKeys.size === 0 || !batchMode) return;

    setBatchProcessing(true);
    setBatchProgress({ current: 0, total: selectedKeys.size });
    try {
      const res = await fetch(`/api/files/${storage.id}/${path}?action=batch`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          operation: batchMode,
          items: Array.from(selectedKeys),
          destPath: batchDestPath,
        }),
      });
      const data = (await res.json()) as { success?: string[]; failed?: { key: string; error: string }[] };
      if (res.ok) {
        const opName = batchMode === "move" ? "移动" : "复制";
        const failedCount = data.failed?.length || 0;
        if (failedCount > 0) {
          alert(`${opName}完成，${failedCount} 个项目${opName}失败`);
        }
      } else {
        alert((data as { error?: string }).error || "操作失败");
      }
      setBatchMode(null);
      setSelectedKeys(new Set());
      loadFiles();
    } catch {
      alert("网络错误");
    } finally {
      setBatchProcessing(false);
      setBatchProgress(null);
    }
  };

  const handleBatchDownload = () => {
    if (selectedKeys.size === 0) return;

    const selectedFiles = objects.filter((obj) => selectedKeys.has(obj.key) && !obj.isDirectory);
    if (selectedFiles.length === 0) {
      alert("选中的项目中没有可下载的文件");
      return;
    }

    if (selectedFiles.length > 10 && !confirm(`即将下载 ${selectedFiles.length} 个文件，是否继续？`)) return;

    // Trigger downloads with small delay to avoid browser blocking
    selectedFiles.forEach((file, index) => {
      setTimeout(() => {
        const link = document.createElement("a");
        link.href = `/api/files/${storage.id}/${file.key}?action=download`;
        link.download = file.name;
        link.style.display = "none";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      }, index * 300);
    });
  };

  const startBatchMoveCopy = async (mode: "move" | "copy") => {
    setBatchMode(mode);
    setBatchDestPath("");
    await loadAllFolders();
  };

  const handleUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const CHUNK_SIZE = chunkSizeMB * 1024 * 1024;

    for (const file of Array.from(files)) {
      try {
        const uploadPath = path ? `${path}/${file.name}` : file.name;

        const canMultipart = supportsMultipart(storage.type);
        if (file.size >= CHUNK_SIZE && canMultipart) {
          await uploadMultipart(file, uploadPath, CHUNK_SIZE);
        } else {
          await uploadSingle(file, uploadPath);
        }
      } catch (err) {
        alert(`上传 ${file.name} 失败: ${err instanceof Error ? err.message : "未知错误"}`);
      }
    }
    setUploadProgress(null);
    loadFiles();
    e.target.value = "";
  };

  const uploadSingle = async (file: File, uploadPath: string) => {
    await new Promise<void>((resolve, reject) => {
      const xhr = new XMLHttpRequest();

      xhr.upload.onprogress = (event) => {
        if (event.lengthComputable) {
          const percent = Math.round((event.loaded / event.total) * 100);
          setUploadProgress({ name: file.name, progress: percent });
        }
      };

      xhr.onload = () => {
        if (xhr.status >= 200 && xhr.status < 300) {
          resolve();
        } else {
          try {
            const data = JSON.parse(xhr.responseText);
            reject(new Error(data.error || "上传失败"));
          } catch {
            reject(new Error("上传失败"));
          }
        }
      };

      xhr.onerror = () => reject(new Error("网络错误"));

      xhr.open("PUT", `/api/files/${storage.id}/${uploadPath}`);
      xhr.setRequestHeader("Content-Type", file.type || "application/octet-stream");
      xhr.send(file);
    });
  };

  const uploadMultipart = async (file: File, uploadPath: string, chunkSize: number) => {
    const totalParts = Math.ceil(file.size / chunkSize);
    const contentType = file.type || "application/octet-stream";
    const CONCURRENT_UPLOADS = 5;

    // Check for existing upload in localStorage (resume support)
    const storageKey = `multipart_${storage.id}_${uploadPath}_${file.size}`;
    const savedState = localStorage.getItem(storageKey);
    let uploadId: string;
    let completedParts: { partNumber: number; etag: string }[] = [];
    let startPart = 0;
    let useDirectUpload = true; // Try direct S3 upload first

    if (savedState) {
      try {
        const parsed = JSON.parse(savedState);
        if (parsed.uploadId && parsed.parts && parsed.fileName === file.name) {
          const shouldResume = confirm(`检测到未完成的上传 "${file.name}"，是否继续？\n已完成 ${parsed.parts.length}/${totalParts} 分片`);
          if (shouldResume) {
            uploadId = parsed.uploadId;
            completedParts = parsed.parts;
            startPart = completedParts.length;
            useDirectUpload = parsed.useDirectUpload ?? true;
          } else {
            try {
              await fetch(`/api/files/${storage.id}/${uploadPath}?action=multipart-abort`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ uploadId: parsed.uploadId }),
              });
            } catch { /* ignore */ }
            localStorage.removeItem(storageKey);
          }
        }
      } catch { /* ignore invalid state */ }
    }

    // Initialize new upload if needed
    if (!uploadId!) {
      setUploadProgress({ name: file.name, progress: 0, currentPart: 0, totalParts, speed: 0, loaded: 0, total: file.size });

      const initRes = await fetch(`/api/files/${storage.id}/${uploadPath}?action=multipart-init`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ contentType, size: file.size, chunkSize }),
      });

      if (!initRes.ok) {
        const data = await initRes.json() as { error?: string };
        throw new Error(data.error || "初始化分片上传失败");
      }

      const initData = await initRes.json() as { uploadId: string };
      uploadId = initData.uploadId;

      localStorage.setItem(storageKey, JSON.stringify({
        uploadId,
        fileName: file.name,
        parts: [],
        useDirectUpload: true,
      }));
    }

    // Speed calculation
    let totalBytesUploaded = startPart * chunkSize;
    const startTime = Date.now();
    const partProgress: Record<number, number> = {};

    const updateProgress = () => {
      const currentBytes = totalBytesUploaded + Object.values(partProgress).reduce((a, b) => a + b, 0);
      const elapsed = (Date.now() - startTime) / 1000;
      const speed = elapsed > 0 ? currentBytes / elapsed : 0;
      const progress = Math.round((currentBytes / file.size) * 100);

      setUploadProgress({
        name: file.name,
        progress: Math.min(progress, 100),
        currentPart: completedParts.length,
        totalParts,
        speed,
        loaded: currentBytes,
        total: file.size,
      });
    };

    updateProgress();

    try {
      const remainingParts = Array.from({ length: totalParts - startPart }, (_, i) => startPart + i + 1);

      // Get signed URLs for direct upload
      let signedUrls: Record<number, string> = {};
      if (useDirectUpload) {
        try {
          const urlsRes = await fetch(`/api/files/${storage.id}/${uploadPath}?action=multipart-urls`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ uploadId, partNumbers: remainingParts }),
          });
          if (urlsRes.ok) {
            const data = await urlsRes.json() as { urls: Record<number, string> };
            signedUrls = data.urls;
          }
        } catch { /* will fallback to proxy */ }
      }

      const uploadQueue = remainingParts.map((partNumber) => ({
        partNumber,
        start: (partNumber - 1) * chunkSize,
        end: Math.min(partNumber * chunkSize, file.size),
      }));

      // Upload part - tries direct S3 first, falls back to Workers proxy
      const uploadPart = async (item: { partNumber: number; start: number; end: number }): Promise<{ partNumber: number; etag: string }> => {
        const chunk = file.slice(item.start, item.end);

        // Try direct S3 upload first
        if (useDirectUpload && signedUrls[item.partNumber]) {
          try {
            const result = await uploadPartDirect(chunk, signedUrls[item.partNumber], item.partNumber);
            return result;
          } catch (e) {
            // CORS or network error - switch to proxy mode
            console.log("Direct upload failed, switching to proxy mode");
            useDirectUpload = false;
            // Update saved state
            localStorage.setItem(storageKey, JSON.stringify({
              uploadId,
              fileName: file.name,
              parts: completedParts,
              useDirectUpload: false,
            }));
          }
        }

        // Fallback: upload through Workers proxy
        return uploadPartProxy(chunk, uploadPath, uploadId, item.partNumber);
      };

      const uploadPartDirect = (chunk: Blob, url: string, partNumber: number): Promise<{ partNumber: number; etag: string }> => {
        return new Promise((resolve, reject) => {
          const xhr = new XMLHttpRequest();

          xhr.upload.onprogress = (event) => {
            if (event.lengthComputable) {
              partProgress[partNumber] = event.loaded;
              updateProgress();
            }
          };

          xhr.onload = () => {
            delete partProgress[partNumber];
            if (xhr.status >= 200 && xhr.status < 300) {
              const etag = xhr.getResponseHeader("ETag")?.replace(/"/g, "") || "";
              totalBytesUploaded += chunk.size;
              resolve({ partNumber, etag });
            } else {
              reject(new Error(`Direct upload failed: ${xhr.status}`));
            }
          };

          xhr.onerror = () => {
            delete partProgress[partNumber];
            reject(new Error("Direct upload network error"));
          };

          xhr.open("PUT", url);
          xhr.send(chunk);
        });
      };

      const uploadPartProxy = (chunk: Blob, path: string, upId: string, partNumber: number): Promise<{ partNumber: number; etag: string }> => {
        return new Promise((resolve, reject) => {
          const xhr = new XMLHttpRequest();

          xhr.upload.onprogress = (event) => {
            if (event.lengthComputable) {
              partProgress[partNumber] = event.loaded;
              updateProgress();
            }
          };

          xhr.onload = () => {
            delete partProgress[partNumber];
            if (xhr.status >= 200 && xhr.status < 300) {
              try {
                const data = JSON.parse(xhr.responseText);
                totalBytesUploaded += chunk.size;
                resolve({ partNumber, etag: data.etag });
              } catch {
                reject(new Error(`解析响应失败: 分片 ${partNumber}`));
              }
            } else {
              try {
                const data = JSON.parse(xhr.responseText);
                reject(new Error(data.error || `分片 ${partNumber} 失败`));
              } catch {
                reject(new Error(`分片 ${partNumber} 失败: ${xhr.status}`));
              }
            }
          };

          xhr.onerror = () => {
            delete partProgress[partNumber];
            reject(new Error(`网络错误: 分片 ${partNumber}`));
          };

          const url = `/api/files/${storage.id}/${path}?action=multipart-upload&uploadId=${encodeURIComponent(upId)}&partNumber=${partNumber}`;
          xhr.open("PUT", url);
          xhr.send(chunk);
        });
      };

      // Process queue with concurrency limit
      let index = 0;

      const runNext = async (): Promise<void> => {
        while (index < uploadQueue.length) {
          const currentIndex = index++;
          const item = uploadQueue[currentIndex];
          const result = await uploadPart(item);
          completedParts.push(result);

          localStorage.setItem(storageKey, JSON.stringify({
            uploadId,
            fileName: file.name,
            parts: completedParts,
            useDirectUpload,
          }));

          updateProgress();
        }
      };

      // Start concurrent uploads (reduce concurrency for proxy mode)
      const concurrency = useDirectUpload ? CONCURRENT_UPLOADS : 3;
      const workers = Array(Math.min(concurrency, uploadQueue.length))
        .fill(null)
        .map(() => runNext());

      await Promise.all(workers);

      // Complete multipart upload
      const completeRes = await fetch(`/api/files/${storage.id}/${uploadPath}?action=multipart-complete`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ uploadId, parts: completedParts }),
      });

      if (!completeRes.ok) {
        const data = await completeRes.json() as { error?: string };
        throw new Error(data.error || "完成分片上传失败");
      }

      localStorage.removeItem(storageKey);
    } catch (err) {
      throw err;
    }
  };

  const handleCreateFolder = async () => {
    if (!newFolderName.trim()) return;

    setCreatingFolder(true);
    try {
      const folderPath = path ? `${path}/${newFolderName.trim()}` : newFolderName.trim();
      const res = await fetch(`/api/files/${storage.id}/${folderPath}?action=mkdir`, {
        method: "POST",
      });

      if (res.ok) {
        setNewFolderName("");
        setShowNewFolderInput(false);
        loadFiles();
      } else {
        const data = (await res.json()) as { error?: string };
        alert(data.error || "创建文件夹失败");
      }
    } catch {
      alert("网络错误");
    } finally {
      setCreatingFolder(false);
    }
  };

  const handleOfflineDownload = async () => {
    if (!offlineUrl.trim()) return;

    setOfflineDownloading(true);
    try {
      const res = await fetch(`/api/files/${storage.id}/${path}?action=fetch`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          url: offlineUrl.trim(),
          filename: offlineFilename.trim() || undefined,
        }),
      });

      const data = await res.json() as { success?: boolean; filename?: string; size?: number; error?: string };

      if (res.ok && data.success) {
        const sizeStr = data.size ? ` (${formatBytes(data.size)})` : "";
        alert(`下载成功: ${data.filename}${sizeStr}`);
        setOfflineUrl("");
        setOfflineFilename("");
        setShowOfflineDownload(false);
        loadFiles();
      } else {
        alert(data.error || "下载失败");
      }
    } catch {
      alert("网络错误");
    } finally {
      setOfflineDownloading(false);
    }
  };

  const breadcrumbs = path ? path.split("/").filter(Boolean) : [];

  const normalizedQuery = searchQuery.trim().toLowerCase();

  // Full-storage search effect with debounce
  useEffect(() => {
    if (!normalizedQuery) {
      setSearchResults(null);
      return;
    }
    const timer = setTimeout(async () => {
      setSearchLoading(true);
      try {
        const res = await fetch(`/api/files/${storage.id}/${encodeURIComponent(path || "")}?action=search&q=${encodeURIComponent(normalizedQuery)}`);
        const data = await res.json();
        setSearchResults(data.objects || []);
      } catch {
        setSearchResults([]);
      } finally {
        setSearchLoading(false);
      }
    }, 300);
    return () => clearTimeout(timer);
  }, [normalizedQuery, storage.id, path]);

  const visibleObjects = normalizedQuery
    ? (searchResults ?? objects.filter((obj) => obj.name.toLowerCase().includes(normalizedQuery)))
    : objects;
  const allVisibleSelected = visibleObjects.length > 0 && visibleObjects.every((obj) => selectedKeys.has(obj.key));

  // Get previewable files for navigation
  const previewableFiles = visibleObjects.filter((obj) => !obj.isDirectory && isPreviewable(obj.name));
  const currentPreviewIndex = previewFile ? previewableFiles.findIndex((f) => f.key === previewFile.key) : -1;

  const handlePreview = (obj: S3Object) => {
    if (isPreviewable(obj.name)) {
      setPreviewFile(obj);
    }
  };

  const handlePrevPreview = () => {
    if (currentPreviewIndex > 0) {
      setPreviewFile(previewableFiles[currentPreviewIndex - 1]);
    }
  };

  const handleNextPreview = () => {
    if (currentPreviewIndex < previewableFiles.length - 1) {
      setPreviewFile(previewableFiles[currentPreviewIndex + 1]);
    }
  };

  // Get file type color class
  const getFileTypeColor = (fileName: string): string => {
    const type = getFileType(fileName);
    switch (type) {
      case 'video': return 'text-purple-500';
      case 'audio': return 'text-pink-500';
      case 'image': return 'text-emerald-500';
      case 'pdf': return 'text-red-500';
      case 'code': return 'text-amber-500';
      case 'markdown': return 'text-blue-400';
      case 'text': return 'text-slate-400';
      default: return 'text-zinc-400 dark:text-zinc-500';
    }
  };

  // SVG file icon component with type-based coloring
  const FileIcon = ({ fileName, size = 18 }: { fileName: string; size?: number }) => {
    const type = getFileType(fileName);
    const colorClass = getFileTypeColor(fileName);
    switch (type) {
      case 'video':
        return <svg className={`shrink-0 ${colorClass}`} width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><polygon points="23 7 16 12 23 17 23 7"/><rect x="1" y="5" width="15" height="14" rx="2" ry="2"/></svg>;
      case 'audio':
        return <svg className={`shrink-0 ${colorClass}`} width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M9 18V5l12-2v13"/><circle cx="6" cy="18" r="3"/><circle cx="18" cy="16" r="3"/></svg>;
      case 'image':
        return <svg className={`shrink-0 ${colorClass}`} width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>;
      case 'pdf':
        return <svg className={`shrink-0 ${colorClass}`} width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="9" y1="15" x2="15" y2="15"/><line x1="9" y1="11" x2="13" y2="11"/></svg>;
      case 'code':
        return <svg className={`shrink-0 ${colorClass}`} width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></svg>;
      case 'markdown':
        return <svg className={`shrink-0 ${colorClass}`} width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M7 13h2l1-2 1 4 1-2h2"/></svg>;
      default:
        return <svg className={`shrink-0 ${colorClass}`} width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>;
    }
  };

  // Check if file is an image for thumbnail
  const isImageFile = (fileName: string) => getFileType(fileName) === 'image';

  return (
    <div className="h-full flex flex-col">
      {/* Toolbar - Row 1: Breadcrumb + Search toggle */}
      <div className="flex items-center justify-between py-2 px-3 md:px-4 border-b border-zinc-200 dark:border-zinc-800 bg-white/50 dark:bg-zinc-900/50 gap-2">
        <div className="flex items-center gap-1 text-sm overflow-x-auto min-w-0 flex-1">
          <button onClick={() => setPath("")} className="text-blue-500 hover:text-blue-400 shrink-0 font-medium">
            {storage.name}
          </button>
          {breadcrumbs.map((part, i) => (
            <span key={i} className="flex items-center shrink-0">
              <svg className="text-zinc-400 dark:text-zinc-600 mx-0.5" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><polyline points="9 18 15 12 9 6"/></svg>
              <button
                onClick={() => navigateTo(breadcrumbs.slice(0, i + 1).join("/"))}
                className="text-blue-500 hover:text-blue-400"
              >
                {part}
              </button>
            </span>
          ))}
          {selectedKeys.size > 0 && (
            <span className="ml-1 text-xs text-blue-500 dark:text-blue-400 font-medium bg-blue-50 dark:bg-blue-900/20 px-2 py-0.5 rounded-full shrink-0">
              {selectedKeys.size}
            </span>
          )}
        </div>
        {/* Search toggle (mobile) / Search input (desktop) */}
        <div className="flex items-center gap-1 shrink-0">
          <div className="relative hidden md:block">
            {searchLoading ? (
              <svg className="absolute left-2 top-1/2 -translate-y-1/2 clist-spinner text-blue-500" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
            ) : (
              <svg className="absolute left-2 top-1/2 -translate-y-1/2 text-zinc-400" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
            )}
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="搜索全部文件..."
              className="w-48 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 pl-7 pr-6 py-1.5 text-xs text-zinc-700 dark:text-zinc-200 rounded-lg focus:border-blue-500 focus:ring-1 focus:ring-blue-500/20 focus:outline-none transition"
            />
            {searchQuery && (
              <button
                onClick={() => setSearchQuery("")}
                className="absolute right-1.5 top-1/2 -translate-y-1/2 text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-300 p-0.5"
              >
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
              </button>
            )}
          </div>
          {path && (
            <button onClick={goUp} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition" title="返回上级">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/></svg>
            </button>
          )}
          <button onClick={loadFiles} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition" title="刷新">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>
          </button>
          {isAdmin && (
            <>
              <button onClick={() => setShowNewFolderInput(true)} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition" title="新建文件夹">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/><line x1="12" y1="11" x2="12" y2="17"/><line x1="9" y1="14" x2="15" y2="14"/></svg>
              </button>
              <button onClick={() => setShowOfflineDownload(true)} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition" title="离线下载">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>
              </button>
            </>
          )}
          {canUpload && (
            <label className={`w-8 h-8 flex items-center justify-center rounded-lg bg-blue-600 hover:bg-blue-500 text-white cursor-pointer transition shadow-sm ${uploadProgress ? 'opacity-50 pointer-events-none' : ''}`}>
              {uploadProgress ? (
                <svg className="clist-spinner" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
              ) : (
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
              )}
              <input type="file" multiple onChange={handleUpload} className="hidden" disabled={!!uploadProgress} />
            </label>
          )}
        </div>
      </div>

      {/* Mobile search bar */}
      <div className="md:hidden px-3 py-1.5 border-b border-zinc-200 dark:border-zinc-800 bg-white/50 dark:bg-zinc-900/50">
        <div className="relative">
          <svg className="absolute left-2.5 top-1/2 -translate-y-1/2 text-zinc-400" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="搜索文件..."
            className="w-full bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 pl-8 pr-8 py-2 text-sm text-zinc-700 dark:text-zinc-200 rounded-lg focus:border-blue-500 focus:ring-1 focus:ring-blue-500/20 focus:outline-none transition"
          />
          {searchQuery && (
            <button
              onClick={() => setSearchQuery("")}
              className="absolute right-2 top-1/2 -translate-y-1/2 text-zinc-400 hover:text-zinc-600 p-1"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </button>
          )}
        </div>
      </div>

      {/* Batch action bar (shown when items selected) */}
      {isAdmin && selectedKeys.size > 0 && (
        <div className="flex items-center gap-1 px-3 md:px-4 py-2 border-b border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-zinc-800/50 overflow-x-auto">
          <button
            onClick={() => startBatchMoveCopy("move")}
            disabled={batchProcessing}
            className="inline-flex items-center gap-1.5 text-xs bg-zinc-700 hover:bg-zinc-600 dark:bg-zinc-600 dark:hover:bg-zinc-500 disabled:opacity-50 text-white px-3 py-1.5 rounded-lg transition shadow-sm whitespace-nowrap"
          >
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><path d="M5 12h14"/><polyline points="12 5 19 12 12 19"/></svg>
            移动
          </button>
          <button
            onClick={() => startBatchMoveCopy("copy")}
            disabled={batchProcessing}
            className="inline-flex items-center gap-1.5 text-xs bg-zinc-700 hover:bg-zinc-600 dark:bg-zinc-600 dark:hover:bg-zinc-500 disabled:opacity-50 text-white px-3 py-1.5 rounded-lg transition shadow-sm whitespace-nowrap"
          >
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
            复制
          </button>
          <button
            onClick={() => {
              if (selectedKeys.size === 1) {
                const obj = objects.find(o => selectedKeys.has(o.key));
                if (obj) startShare(obj);
              } else {
                alert("分享功能仅支持单个文件，请只选择一个文件");
              }
            }}
            className="inline-flex items-center gap-1.5 text-xs bg-zinc-700 hover:bg-zinc-600 dark:bg-zinc-600 dark:hover:bg-zinc-500 text-white px-3 py-1.5 rounded-lg transition shadow-sm whitespace-nowrap"
          >
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><line x1="8.59" y1="13.51" x2="15.42" y2="17.49"/><line x1="15.41" y1="6.51" x2="8.59" y2="10.49"/></svg>
            分享
          </button>
          {canDownload && (
            <button
              onClick={handleBatchDownload}
              className="inline-flex items-center gap-1.5 text-xs bg-zinc-700 hover:bg-zinc-600 dark:bg-zinc-600 dark:hover:bg-zinc-500 text-white px-3 py-1.5 rounded-lg transition shadow-sm whitespace-nowrap"
            >
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              下载
            </button>
          )}
          <button
            onClick={handleBatchDelete}
            disabled={deleting}
            className="inline-flex items-center gap-1.5 text-xs bg-red-600 hover:bg-red-500 disabled:opacity-50 text-white px-3 py-1.5 rounded-lg transition shadow-sm whitespace-nowrap"
          >
            {deleting ? (
              <svg className="clist-spinner" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
            ) : (
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
            )}
            删除
          </button>
          <div className="flex-1" />
        </div>
      )}

      {/* New Folder Input */}
      {showNewFolderInput && (
        <div className="px-4 py-2 border-b border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-zinc-800/50">
          <div className="flex items-center gap-2">
            <span className="text-xs text-zinc-500 font-mono">新建文件夹:</span>
            <input
              type="text"
              value={newFolderName}
              onChange={(e) => setNewFolderName(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") handleCreateFolder();
                if (e.key === "Escape") {
                  setShowNewFolderInput(false);
                  setNewFolderName("");
                }
              }}
              placeholder="输入文件夹名称"
              className="flex-1 bg-white dark:bg-zinc-900 border border-zinc-300 dark:border-zinc-700 px-2 py-1 text-sm font-mono text-zinc-900 dark:text-zinc-100 rounded focus:border-blue-500 focus:outline-none"
              autoFocus
              disabled={creatingFolder}
            />
            <button
              onClick={handleCreateFolder}
              disabled={creatingFolder || !newFolderName.trim()}
              className="text-xs bg-blue-600 hover:bg-blue-500 disabled:opacity-50 text-white px-3 py-1 font-mono rounded"
            >
              {creatingFolder ? "创建中..." : "创建"}
            </button>
            <button
              onClick={() => {
                setShowNewFolderInput(false);
                setNewFolderName("");
              }}
              className="text-xs text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 font-mono px-2 py-1"
            >
              取消
            </button>
          </div>
        </div>
      )}

      {/* Offline Download Input */}
      {showOfflineDownload && (
        <div className="px-4 py-2 border-b border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-zinc-800/50">
          <div className="flex flex-col gap-2">
            <div className="flex items-center gap-2">
              <span className="text-xs text-zinc-500 font-mono shrink-0">链接地址:</span>
              <input
                type="url"
                value={offlineUrl}
                onChange={(e) => setOfflineUrl(e.target.value)}
                placeholder="https://example.com/file.zip"
                className="flex-1 bg-white dark:bg-zinc-900 border border-zinc-300 dark:border-zinc-700 px-2 py-1 text-sm font-mono text-zinc-900 dark:text-zinc-100 rounded focus:border-blue-500 focus:outline-none"
                autoFocus
                disabled={offlineDownloading}
              />
            </div>
            <div className="flex items-center gap-2">
              <span className="text-xs text-zinc-500 font-mono shrink-0">文件名称:</span>
              <input
                type="text"
                value={offlineFilename}
                onChange={(e) => setOfflineFilename(e.target.value)}
                placeholder="可选，留空自动识别"
                className="flex-1 bg-white dark:bg-zinc-900 border border-zinc-300 dark:border-zinc-700 px-2 py-1 text-sm font-mono text-zinc-900 dark:text-zinc-100 rounded focus:border-blue-500 focus:outline-none"
                disabled={offlineDownloading}
                onKeyDown={(e) => {
                  if (e.key === "Enter") handleOfflineDownload();
                  if (e.key === "Escape") {
                    setShowOfflineDownload(false);
                    setOfflineUrl("");
                    setOfflineFilename("");
                  }
                }}
              />
              <button
                onClick={handleOfflineDownload}
                disabled={offlineDownloading || !offlineUrl.trim()}
                className="text-xs bg-blue-600 hover:bg-blue-500 disabled:opacity-50 text-white px-3 py-1 font-mono rounded whitespace-nowrap"
              >
                {offlineDownloading ? "下载中..." : "开始下载"}
              </button>
              <button
                onClick={() => {
                  setShowOfflineDownload(false);
                  setOfflineUrl("");
                  setOfflineFilename("");
                }}
                disabled={offlineDownloading}
                className="text-xs text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 font-mono px-2 py-1"
              >
                取消
              </button>
            </div>
            <p className="text-xs text-zinc-400 dark:text-zinc-600 font-mono">
              提示: 文件将下载到当前目录，大文件可能需要较长时间
            </p>
          </div>
        </div>
      )}

      {/* Upload Progress */}
      {uploadProgress && (
        <div className="px-4 py-2 border-b border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-zinc-800/50">
          <div className="flex items-center gap-3">
            <span className="text-xs text-zinc-600 dark:text-zinc-400 font-mono truncate flex-1">
              正在上传: {uploadProgress.name}
              {uploadProgress.totalParts && (
                <span className="text-zinc-400 dark:text-zinc-500 ml-1">
                  ({uploadProgress.currentPart}/{uploadProgress.totalParts} 分片)
                </span>
              )}
            </span>
            {uploadProgress.speed !== undefined && uploadProgress.speed > 0 && (
              <span className="text-xs text-blue-500 font-mono shrink-0">
                {formatSpeed(uploadProgress.speed)}
              </span>
            )}
            <span className="text-xs text-zinc-500 font-mono w-12 text-right">
              {uploadProgress.progress}%
            </span>
          </div>
          <div className="mt-1 h-1 bg-zinc-200 dark:bg-zinc-700 rounded-full overflow-hidden">
            <div
              className="h-full bg-blue-500 transition-all duration-150"
              style={{ width: `${uploadProgress.progress}%` }}
            />
          </div>
          {uploadProgress.loaded !== undefined && uploadProgress.total !== undefined && (
            <div className="mt-1 text-xs text-zinc-400 dark:text-zinc-500 font-mono">
              {formatBytes(uploadProgress.loaded)} / {formatBytes(uploadProgress.total)}
            </div>
          )}
        </div>
      )}

      {/* Content */}
      <div className="flex-1 overflow-auto">
        {loading ? (
          <div className="flex flex-col items-center justify-center h-40 gap-3">
            <svg className="clist-spinner text-blue-500" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
            <span className="text-sm text-zinc-500 dark:text-zinc-400">加载中...</span>
          </div>
        ) : error ? (
          <div className="flex flex-col items-center justify-center h-40 gap-3">
            <div className="w-12 h-12 rounded-full bg-red-50 dark:bg-red-900/20 flex items-center justify-center">
              <svg className="text-red-400" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>
            </div>
            <span className="text-sm text-red-500 dark:text-red-400">{error}</span>
          </div>
        ) : objects.length === 0 ? (
          <div className="flex flex-col items-center justify-center h-40 gap-3">
            <div className="w-14 h-14 rounded-2xl bg-zinc-100 dark:bg-zinc-800 flex items-center justify-center">
              <svg className="text-zinc-300 dark:text-zinc-600" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
                <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
              </svg>
            </div>
            <div className="text-center">
              <p className="text-sm font-medium text-zinc-500 dark:text-zinc-400">空目录</p>
              <p className="text-xs text-zinc-400 dark:text-zinc-600 mt-0.5">此目录中没有文件</p>
            </div>
          </div>
        ) : (
          <>
            {/* Desktop Table View */}
            <table className="w-full text-sm hidden md:table">
              <thead className="text-xs text-zinc-500 border-b border-zinc-200 dark:border-zinc-800 sticky top-0 bg-zinc-50 dark:bg-zinc-900">
                <tr>
                  {isAdmin && (
                    <th className="py-2.5 px-3 w-10 text-left">
                      <input
                        type="checkbox"
                        checked={allVisibleSelected}
                        onChange={toggleSelectAll}
                        className="w-4 h-4 rounded border-zinc-300 dark:border-zinc-600 bg-white dark:bg-zinc-800"
                      />
                    </th>
                  )}
                  <th className="text-left py-2.5 px-4 font-medium">名称</th>
                  <th className="text-right py-2.5 px-4 font-medium w-24">大小</th>
                  <th className="text-right py-2.5 px-4 font-medium w-44">修改时间</th>
                  <th className="text-right py-2.5 px-4 font-medium w-28">操作</th>
                </tr>
              </thead>
              <tbody>
                {visibleObjects.length === 0 ? (
                  <tr>
                    <td
                      colSpan={isAdmin ? 5 : 4}
                      className="py-8 text-center text-zinc-400 dark:text-zinc-600"
                    >
                      没有匹配的文件
                    </td>
                  </tr>
                ) : visibleObjects.map((obj, idx) => (
                  <tr
                    key={obj.key}
                    className={`border-b border-zinc-100 dark:border-zinc-800/50 clist-file-row transition-colors ${
                      selectedKeys.has(obj.key) ? "bg-blue-50/60 dark:bg-blue-900/15" : idx % 2 === 0 ? "bg-white dark:bg-transparent" : "bg-zinc-50/50 dark:bg-zinc-800/10"
                    }`}
                  >
                    {isAdmin && (
                      <td className="py-2.5 px-3">
                        <input
                          type="checkbox"
                          checked={selectedKeys.has(obj.key)}
                          onChange={() => toggleSelect(obj.key)}
                          className="w-4 h-4 rounded border-zinc-300 dark:border-zinc-600 bg-white dark:bg-zinc-800"
                        />
                      </td>
                    )}
                    <td className="py-2.5 px-4">
                      {obj.isDirectory ? (
                        <button
                          onClick={() => navigateTo(obj.key)}
                          className="flex items-center gap-2 text-blue-500 hover:text-blue-400 transition"
                        >
                          <svg className="text-yellow-500 shrink-0" width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="M2 6a2 2 0 0 1 2-2h5l2 2h9a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V6z"/></svg>
                          <span className="truncate">{obj.name}</span>
                        </button>
                      ) : (
                        <div className="flex items-center gap-3 max-w-md">
                          {/* Thumbnail for images, icon for others */}
                          {isImageFile(obj.name) ? (
                            <div className="shrink-0 w-10 h-10 rounded-lg overflow-hidden bg-zinc-100 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700/50 clist-thumb">
                              <img
                                src={`/api/files/${storage.id}/${obj.key}?action=download`}
                                alt=""
                                className="w-full h-full object-cover"
                                loading="lazy"
                                onError={(e) => { (e.target as HTMLImageElement).style.display = 'none'; }}
                              />
                            </div>
                          ) : (
                            <FileIcon fileName={obj.name} />
                          )}
                          {isPreviewable(obj.name) ? (
                            <button
                              onClick={() => handlePreview(obj)}
                              className="text-zinc-700 dark:text-zinc-300 hover:text-blue-500 dark:hover:text-blue-400 transition truncate text-left"
                            >
                              {obj.name}
                              {normalizedQuery && obj.key !== obj.name && (
                                <span className="block text-[10px] text-zinc-400 dark:text-zinc-600 font-normal truncate">{obj.key.replace(/[^/]+$/, "").replace(/\/$/, "")}</span>
                              )}
                            </button>
                          ) : (
                            <span className="text-zinc-700 dark:text-zinc-300 truncate">
                              {obj.name}
                              {normalizedQuery && obj.key !== obj.name && (
                                <span className="block text-[10px] text-zinc-400 dark:text-zinc-600 truncate">{obj.key.replace(/[^/]+$/, "").replace(/\/$/, "")}</span>
                              )}
                            </span>
                          )}
                        </div>
                      )}
                    </td>
                    <td className="py-2.5 px-4 text-right text-zinc-500 whitespace-nowrap">
                      {obj.isDirectory ? "-" : formatBytes(obj.size)}
                    </td>
                    <td className="py-2.5 px-4 text-right text-zinc-500 whitespace-nowrap text-xs">
                      {formatDate(obj.lastModified)}
                    </td>
                    <td className="py-2.5 px-4 text-right">
                      {obj.isDirectory ? (
                        isAdmin && (
                          <div className="flex items-center justify-end gap-0.5">
                            <button
                              onClick={() => startShare(obj)}
                              data-tooltip="分享"
                              className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20 transition"
                            >
                              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><line x1="8.59" y1="13.51" x2="15.42" y2="17.49"/><line x1="15.41" y1="6.51" x2="8.59" y2="10.49"/></svg>
                            </button>
                            <button
                              onClick={() => startRename(obj)}
                              data-tooltip="重命名"
                              className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20 transition"
                            >
                              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                            <button
                              onClick={() => startMove(obj)}
                              data-tooltip="移动"
                              className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20 transition"
                            >
                              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><path d="M5 12h14"/><polyline points="12 5 19 12 12 19"/></svg>
                            </button>
                            <button
                              onClick={() => deleteFolder(obj.key, obj.name)}
                              data-tooltip="删除"
                              className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 transition"
                            >
                              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                            </button>
                          </div>
                        )
                      ) : (
                        <div className="flex items-center justify-end gap-0.5">
                          {canDownload && isPreviewable(obj.name) && (
                            <button
                              onClick={() => handlePreview(obj)}
                              data-tooltip="预览"
                              className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20 transition"
                            >
                              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><polygon points="5 3 19 12 5 21 5 3"/></svg>
                            </button>
                          )}
                          {canDownload && (
                            <button
                              onClick={() => downloadFile(obj.key)}
                              data-tooltip="下载"
                              className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20 transition"
                            >
                              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                            </button>
                          )}
                          {isAdmin && (
                            <>
                              <button
                                onClick={() => startShare(obj)}
                                data-tooltip="分享"
                                className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20 transition"
                              >
                                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><line x1="8.59" y1="13.51" x2="15.42" y2="17.49"/><line x1="15.41" y1="6.51" x2="8.59" y2="10.49"/></svg>
                              </button>
                              <button
                                onClick={() => startRename(obj)}
                                data-tooltip="重命名"
                                className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20 transition"
                              >
                                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                              </button>
                              <button
                                onClick={() => startMove(obj)}
                                data-tooltip="移动"
                                className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20 transition"
                              >
                                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><path d="M5 12h14"/><polyline points="12 5 19 12 12 19"/></svg>
                              </button>
                              <button
                                onClick={() => deleteFile(obj.key)}
                                data-tooltip="删除"
                                className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 transition"
                              >
                                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                              </button>
                            </>
                          )}
                        </div>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

            {/* Mobile Card/List View */}
            <div className="md:hidden divide-y divide-zinc-100 dark:divide-zinc-800/50">
              {visibleObjects.length === 0 ? (
                <div className="py-8 text-center text-zinc-400 dark:text-zinc-600 text-sm">
                  没有匹配的文件
                </div>
              ) : visibleObjects.map((obj) => (
                <div
                  key={obj.key}
                  className={`flex items-center gap-3 px-4 py-3 ${
                    selectedKeys.has(obj.key) ? "bg-blue-50/60 dark:bg-blue-900/15" : ""
                  }`}
                >
                  {isAdmin && (
                    <input
                      type="checkbox"
                      checked={selectedKeys.has(obj.key)}
                      onChange={() => toggleSelect(obj.key)}
                      className="w-5 h-5 rounded border-zinc-300 dark:border-zinc-600 bg-white dark:bg-zinc-800 shrink-0"
                    />
                  )}
                  <div className="flex-1 min-w-0">
                    {obj.isDirectory ? (
                      <button
                        onClick={() => navigateTo(obj.key)}
                        className="flex items-center gap-2 text-blue-500 hover:text-blue-400 w-full text-left"
                      >
                        <svg className="text-yellow-500 shrink-0" width="20" height="20" viewBox="0 0 24 24" fill="currentColor"><path d="M2 6a2 2 0 0 1 2-2h5l2 2h9a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V6z"/></svg>
                        <span className="line-clamp-2 font-medium leading-snug">{obj.name}</span>
                      </button>
                    ) : (
                      <button
                        onClick={() => isPreviewable(obj.name) ? handlePreview(obj) : undefined}
                        className="flex items-center gap-3 w-full text-left"
                      >
                        {/* Thumbnail for images, icon for others */}
                        {isImageFile(obj.name) ? (
                          <div className="shrink-0 w-11 h-11 rounded-lg overflow-hidden bg-zinc-100 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700/50 clist-thumb">
                            <img
                              src={`/api/files/${storage.id}/${obj.key}?action=download`}
                              alt=""
                              className="w-full h-full object-cover"
                              loading="lazy"
                              onError={(e) => { (e.target as HTMLImageElement).style.display = 'none'; }}
                            />
                          </div>
                        ) : (
                          <FileIcon fileName={obj.name} size={20} />
                        )}
                        <div className="min-w-0">
                          <div className="line-clamp-2 text-zinc-700 dark:text-zinc-300 leading-snug">{obj.name}</div>
                          <div className="text-xs text-zinc-400 dark:text-zinc-600 mt-0.5">
                            {normalizedQuery && obj.key !== obj.name && (
                              <span className="text-blue-500 dark:text-blue-400">{obj.key.replace(/[^/]+$/, "").replace(/\/$/, "")} · </span>
                            )}
                            {formatBytes(obj.size)} · {formatDate(obj.lastModified)}
                          </div>
                        </div>
                      </button>
                    )}
                  </div>
                  <div className="flex items-center gap-1 shrink-0">
                    {canDownload && !obj.isDirectory && (
                      <button
                        onClick={() => downloadFile(obj.key)}
                        className="w-10 h-10 flex items-center justify-center rounded-lg text-zinc-400 hover:text-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20 active:bg-blue-100 dark:active:bg-blue-900/30 transition"
                      >
                        <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                      </button>
                    )}
                    {isAdmin && (
                      <button
                        onClick={() => obj.isDirectory ? deleteFolder(obj.key, obj.name) : deleteFile(obj.key)}
                        className="w-10 h-10 flex items-center justify-center rounded-lg text-zinc-400 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 active:bg-red-100 dark:active:bg-red-900/30 transition"
                      >
                        <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                      </button>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </>
        )}
      </div>

      {/* File Preview Modal */}
      {previewFile && (
        <FilePreview
          storageId={storage.id}
          fileKey={previewFile.key}
          fileName={previewFile.name}
          onClose={() => setPreviewFile(null)}
          onPrev={handlePrevPreview}
          onNext={handleNextPreview}
          hasPrev={currentPreviewIndex > 0}
          hasNext={currentPreviewIndex < previewableFiles.length - 1}
        />
      )}

      {/* Rename Modal */}
      {renameTarget && (
        <div className="fixed inset-0 bg-black/50 dark:bg-black/70 backdrop-blur-sm flex items-center justify-center z-50 p-4" onClick={() => setRenameTarget(null)}>
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700/50 w-full max-w-sm rounded-xl shadow-2xl clist-modal-enter" onClick={e => e.stopPropagation()}>
            <div className="px-5 py-4 border-b border-zinc-200 dark:border-zinc-700/50 flex items-center justify-between">
              <span className="text-zinc-900 dark:text-zinc-100 font-medium text-sm">重命名</span>
              <button onClick={() => setRenameTarget(null)} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
              </button>
            </div>
            <div className="p-5 space-y-4">
              <div>
                <label className="block text-xs text-zinc-500 mb-1.5 font-medium">新名称</label>
                <input
                  type="text"
                  value={renameValue}
                  onChange={(e) => setRenameValue(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && handleRename()}
                  className="w-full bg-zinc-50 dark:bg-zinc-800/80 border border-zinc-200 dark:border-zinc-700 px-3 py-2.5 text-zinc-900 dark:text-zinc-100 text-sm rounded-lg focus:border-blue-500 focus:ring-1 focus:ring-blue-500/20 focus:outline-none transition"
                  autoFocus
                />
              </div>
              <div className="flex gap-2 pt-1">
                <button
                  onClick={() => setRenameTarget(null)}
                  className="flex-1 px-4 py-2.5 border border-zinc-200 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400 text-sm hover:bg-zinc-50 dark:hover:bg-zinc-800 rounded-lg transition"
                >
                  取消
                </button>
                <button
                  onClick={handleRename}
                  disabled={renaming || !renameValue.trim()}
                  className="flex-1 px-4 py-2.5 bg-blue-600 hover:bg-blue-500 text-white text-sm font-medium disabled:opacity-50 rounded-lg shadow-sm transition"
                >
                  {renaming ? "处理中..." : "确定"}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Share Modal */}
      {shareTarget && (
        <div className="fixed inset-0 bg-black/50 dark:bg-black/70 backdrop-blur-sm flex items-center justify-center z-50 p-4" onClick={() => setShareTarget(null)}>
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700/50 w-full max-w-sm rounded-xl shadow-2xl clist-modal-enter" onClick={e => e.stopPropagation()}>
            <div className="px-5 py-4 border-b border-zinc-200 dark:border-zinc-700/50 flex items-center justify-between">
              <span className="text-zinc-900 dark:text-zinc-100 font-medium text-sm">生成分享链接</span>
              <button onClick={() => setShareTarget(null)} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
              </button>
            </div>
            <div className="p-5 space-y-4">
              <div className="text-xs text-zinc-500 mb-2 font-medium">
                分享: {shareTarget.name}
              </div>

              {!shareUrl ? (
                <>
                  <div>
                    <label className="block text-xs text-zinc-500 mb-1 font-mono">过期时间</label>
                    <select
                      value={shareExpireHours}
                      onChange={(e) => setShareExpireHours(parseInt(e.target.value, 10))}
                      className="w-full bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 px-3 py-2 text-zinc-900 dark:text-zinc-100 font-mono text-sm focus:border-blue-500 focus:outline-none rounded"
                    >
                      <option value={0}>永不过期</option>
                      <option value={1}>1 小时</option>
                      <option value={24}>1 天</option>
                      <option value={168}>1 周</option>
                      <option value={720}>1 月</option>
                    </select>
                  </div>
                  <div className="flex gap-2">
                    <button
                      onClick={() => setShareTarget(null)}
                      className="flex-1 px-3 py-2 border border-zinc-200 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400 text-sm hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded"
                    >
                      取消
                    </button>
                    <button
                      onClick={handleCreateShare}
                      disabled={creatingShare}
                      className="flex-1 px-3 py-2 bg-blue-500 text-white text-sm hover:bg-blue-600 disabled:opacity-50 rounded"
                    >
                      {creatingShare ? "生成中..." : "生成链接"}
                    </button>
                  </div>
                </>
              ) : (
                <>
                  <div className="space-y-2">
                    <div>
                      <label className="block text-xs text-zinc-500 mb-1 font-mono">分享令牌</label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          value={shareToken}
                          readOnly
                          className="flex-1 bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 px-3 py-2 text-zinc-900 dark:text-zinc-100 font-mono text-xs rounded"
                        />
                        <button
                          onClick={() => copyToClipboard(shareToken)}
                          className="px-3 py-2 bg-zinc-200 dark:bg-zinc-700 hover:bg-zinc-300 dark:hover:bg-zinc-600 text-sm rounded text-zinc-900 dark:text-zinc-100"
                        >
                          复制
                        </button>
                      </div>
                    </div>
                    <div>
                      <label className="block text-xs text-zinc-500 mb-1 font-mono">分享链接</label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          value={shareUrl}
                          readOnly
                          className="flex-1 bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 px-3 py-2 text-zinc-900 dark:text-zinc-100 font-mono text-xs rounded overflow-hidden"
                        />
                        <button
                          onClick={() => copyToClipboard(shareUrl)}
                          className="px-3 py-2 bg-zinc-200 dark:bg-zinc-700 hover:bg-zinc-300 dark:hover:bg-zinc-600 text-sm rounded text-zinc-900 dark:text-zinc-100"
                        >
                          复制
                        </button>
                      </div>
                    </div>
                  </div>
                  <div className="flex gap-2">
                    <button
                      onClick={() => setShareTarget(null)}
                      className="flex-1 px-3 py-2 bg-blue-500 text-white text-sm hover:bg-blue-600 rounded"
                    >
                      完成
                    </button>
                  </div>
                </>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Move Modal */}
      {moveTarget && (
        <div className="fixed inset-0 bg-black/50 dark:bg-black/70 backdrop-blur-sm flex items-center justify-center z-50 p-4" onClick={() => setMoveTarget(null)}>
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700/50 w-full max-w-sm rounded-xl shadow-2xl clist-modal-enter" onClick={e => e.stopPropagation()}>
            <div className="px-5 py-4 border-b border-zinc-200 dark:border-zinc-700/50 flex items-center justify-between">
              <span className="text-zinc-900 dark:text-zinc-100 font-medium text-sm">移动到</span>
              <button onClick={() => setMoveTarget(null)} className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
              </button>
            </div>
            <div className="p-5 space-y-4">
              <div className="text-xs text-zinc-500 mb-2 font-medium">
                移动: {moveTarget.name}
              </div>
              <div>
                <label className="block text-xs text-zinc-500 mb-1.5 font-medium">目标文件夹</label>
                <select
                  value={moveDestPath}
                  onChange={(e) => setMoveDestPath(e.target.value)}
                  className="w-full bg-zinc-50 dark:bg-zinc-800/80 border border-zinc-200 dark:border-zinc-700 px-3 py-2.5 text-zinc-900 dark:text-zinc-100 text-sm rounded-lg focus:border-blue-500 focus:ring-1 focus:ring-blue-500/20 focus:outline-none transition"
                >
                  {allFolders.map((folder) => (
                    <option key={folder} value={folder}>
                      {folder === "" ? "/ (根目录)" : "/" + folder}
                    </option>
                  ))}
                </select>
              </div>
              <div className="flex gap-2 pt-1">
                <button
                  onClick={() => setMoveTarget(null)}
                  className="flex-1 px-4 py-2.5 border border-zinc-200 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400 text-sm hover:bg-zinc-50 dark:hover:bg-zinc-800 rounded-lg transition"
                >
                  取消
                </button>
                <button
                  onClick={handleMove}
                  disabled={moving}
                  className="flex-1 px-4 py-2.5 bg-blue-600 hover:bg-blue-500 text-white text-sm font-medium disabled:opacity-50 rounded-lg shadow-sm transition"
                >
                  {moving ? "处理中..." : "确定"}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Batch Move/Copy Modal */}
      {batchMode && (
        <div className="fixed inset-0 bg-black/50 dark:bg-black/70 backdrop-blur-sm flex items-center justify-center z-50 p-4" onClick={() => !batchProcessing && setBatchMode(null)}>
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700/50 w-full max-w-sm rounded-xl shadow-2xl clist-modal-enter" onClick={e => e.stopPropagation()}>
            <div className="px-5 py-4 border-b border-zinc-200 dark:border-zinc-700/50 flex items-center justify-between">
              <span className="text-zinc-900 dark:text-zinc-100 font-medium text-sm">
                批量{batchMode === "move" ? "移动" : "复制"}到
              </span>
              <button
                onClick={() => !batchProcessing && setBatchMode(null)}
                className="w-8 h-8 flex items-center justify-center rounded-lg text-zinc-400 dark:text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition"
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
              </button>
            </div>
            <div className="p-5 space-y-4">
              <div className="text-xs text-zinc-500 mb-2 font-medium">
                {batchMode === "move" ? "移动" : "复制"}: {selectedKeys.size} 个项目
              </div>
              {batchProgress && (
                <div className="flex items-center gap-2 text-xs text-blue-500 font-medium bg-blue-50 dark:bg-blue-900/20 px-3 py-2 rounded-lg">
                  <svg className="clist-spinner" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
                  处理中... ({batchProgress.current}/{batchProgress.total})
                </div>
              )}
              <div>
                <label className="block text-xs text-zinc-500 mb-1.5 font-medium">目标文件夹</label>
                <select
                  value={batchDestPath}
                  onChange={(e) => setBatchDestPath(e.target.value)}
                  disabled={batchProcessing}
                  className="w-full bg-zinc-50 dark:bg-zinc-800/80 border border-zinc-200 dark:border-zinc-700 px-3 py-2.5 text-zinc-900 dark:text-zinc-100 text-sm rounded-lg focus:border-blue-500 focus:ring-1 focus:ring-blue-500/20 focus:outline-none transition disabled:opacity-50"
                >
                  {allFolders.map((folder) => (
                    <option key={folder} value={folder}>
                      {folder === "" ? "/ (根目录)" : "/" + folder}
                    </option>
                  ))}
                </select>
              </div>
              <div className="flex gap-2">
                <button
                  onClick={() => setBatchMode(null)}
                  disabled={batchProcessing}
                  className="flex-1 px-3 py-2 border border-zinc-200 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400 text-sm hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded disabled:opacity-50"
                >
                  取消
                </button>
                <button
                  onClick={handleBatchMoveCopy}
                  disabled={batchProcessing}
                  className="flex-1 px-3 py-2 bg-blue-500 text-white text-sm hover:bg-blue-600 disabled:opacity-50 rounded"
                >
                  {batchProcessing ? "处理中..." : "确定"}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default function Home({ loaderData }: Route.ComponentProps) {
  const [isAdmin, setIsAdmin] = useState(loaderData.isAdmin);
  const [storages, setStorages] = useState<StorageInfo[]>(loaderData.storages);
  const [selectedStorage, setSelectedStorage] = useState<StorageInfo | null>(null);
  const [showLogin, setShowLogin] = useState(false);
  const [showStorageForm, setShowStorageForm] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showAnnouncement, setShowAnnouncement] = useState(false);
  const [showChangelog, setShowChangelog] = useState(false);
  const [showStats, setShowStats] = useState(false);
  const [statsStorage, setStatsStorage] = useState<StorageInfo | null>(null);
  const [editingStorage, setEditingStorage] = useState<StorageInfo | null>(null);
  const [isDark, setIsDark] = useState(false);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  const siteTitle = loaderData.siteTitle || "CList";
  const siteAnnouncement = loaderData.siteAnnouncement || "";
  const chunkSizeMB = loaderData.chunkSizeMB || 50;
  const webdavEnabled = loaderData.webdavEnabled || false;

  useEffect(() => {
    const saved = localStorage.getItem("theme");
    if (saved === "dark") {
      setIsDark(true);
      document.documentElement.classList.add("dark");
    } else {
      document.documentElement.classList.remove("dark");
    }

    // Show announcement on first visit (per session)
    if (siteAnnouncement) {
      const announcementShown = sessionStorage.getItem("announcement_shown");
      if (!announcementShown) {
        setShowAnnouncement(true);
        sessionStorage.setItem("announcement_shown", "true");
      }
    }
  }, [siteAnnouncement]);

  const toggleTheme = useCallback((event: React.MouseEvent) => {
    const newIsDark = !isDark;

    const changeTheme = () => {
      setIsDark(newIsDark);
      if (newIsDark) {
        document.documentElement.classList.add("dark");
        localStorage.setItem("theme", "dark");
      } else {
        document.documentElement.classList.remove("dark");
        localStorage.setItem("theme", "light");
      }
    };

    if (!document.startViewTransition) {
      changeTheme();
      return;
    }

    const x = event.clientX;
    const y = event.clientY;
    const endRadius = Math.hypot(
      Math.max(x, window.innerWidth - x),
      Math.max(y, window.innerHeight - y)
    );

    const transition = document.startViewTransition(() => {
      changeTheme();
    });

    transition.ready.then(() => {
      const clipPath = [
        `circle(0px at ${x}px ${y}px)`,
        `circle(${endRadius}px at ${x}px ${y}px)`,
      ];
      document.documentElement.animate(
        { clipPath: isDark ? clipPath : clipPath.reverse() },
        {
          duration: 400,
          easing: "ease-in-out",
          pseudoElement: isDark
            ? "::view-transition-new(root)"
            : "::view-transition-old(root)",
        }
      );
    });
  }, [isDark]);

  const refreshStorages = async () => {
    try {
      const res = await fetch("/api/storages");
      if (res.ok) {
        const data = (await res.json()) as { storages: StorageInfo[]; isAdmin: boolean };
        setStorages(data.storages);
        setIsAdmin(data.isAdmin);
      }
    } catch { /* ignore */ }
  };

  const handleLogout = async () => {
    try {
      await fetch("/api/storages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "logout" }),
      });
      setIsAdmin(false);
      setSelectedStorage(null);
      refreshStorages();
    } catch { /* ignore */ }
  };

  const handleDeleteStorage = async (s: StorageInfo) => {
    if (!confirm(`删除存储 "${s.name}"?`)) return;
    try {
      const res = await fetch(`/api/storages?id=${s.id}`, { method: "DELETE" });
      if (res.ok) {
        if (selectedStorage?.id === s.id) setSelectedStorage(null);
        refreshStorages();
      }
    } catch { /* ignore */ }
  };

  return (
    <div className="h-screen overflow-hidden bg-zinc-100 dark:bg-zinc-950 text-zinc-900 dark:text-zinc-100 transition-colors flex flex-col">
      {/* Header */}
      <header className="border-b border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 shrink-0 shadow-sm">
        <div className="px-3 md:px-5 py-2.5 flex items-center justify-between gap-2 md:gap-4">
          {/* Left: Hamburger + Logo */}
          <div className="flex items-center gap-2 shrink-0">
            <button
              onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
              className="md:hidden flex items-center justify-center w-9 h-9 rounded-lg text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition"
              title="菜单"
            >
              <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round">
                {sidebarCollapsed ? (
                  <>
                    <line x1="3" y1="5" x2="17" y2="5" />
                    <line x1="3" y1="10" x2="17" y2="10" />
                    <line x1="3" y1="15" x2="17" y2="15" />
                  </>
                ) : (
                  <>
                    <line x1="5" y1="5" x2="15" y2="15" />
                    <line x1="15" y1="5" x2="5" y2="15" />
                  </>
                )}
              </svg>
            </button>
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 rounded-lg bg-gradient-to-br from-blue-500 to-blue-600 flex items-center justify-center shadow-sm">
                <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                  <path d="M2 4.5C2 3.67 2.67 3 3.5 3h9c.83 0 1.5.67 1.5 1.5v7c0 .83-.67 1.5-1.5 1.5h-9C2.67 13 2 12.33 2 11.5v-7z" stroke="white" strokeWidth="1.2"/>
                  <path d="M5 7h6M5 9.5h4" stroke="white" strokeWidth="1.2" strokeLinecap="round"/>
                </svg>
              </div>
              <span className="text-base font-bold tracking-tight hidden sm:inline">CList</span>
            </div>
          </div>

          {/* Center: Site title */}
          <div className="flex-1 text-center min-w-0">
            <span className="text-xs sm:text-sm font-medium text-zinc-500 dark:text-zinc-400 truncate block">{siteTitle}</span>
          </div>

          {/* Right: Actions */}
          <div className="flex items-center gap-1 shrink-0">
            <button
              onClick={toggleTheme}
              data-tooltip={isDark ? "切换亮色" : "切换暗色"}
              className="flex items-center justify-center w-9 h-9 rounded-lg text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition"
            >
              {isDark ? (
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>
                </svg>
              ) : (
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>
                </svg>
              )}
            </button>
            <button
              onClick={() => setShowSettings(true)}
              data-tooltip="设置"
              className="flex items-center justify-center w-9 h-9 rounded-lg text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition"
            >
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/>
              </svg>
            </button>
            {isAdmin ? (
              <>
                <span className="hidden sm:inline-flex items-center gap-1.5 text-xs text-green-600 dark:text-green-500 font-medium px-2 py-1 rounded-full bg-green-50 dark:bg-green-900/20">
                  <span className="w-1.5 h-1.5 rounded-full bg-green-500"></span>
                  管理员
                </span>
                <button
                  onClick={handleLogout}
                  data-tooltip="登出"
                  className="flex items-center justify-center w-9 h-9 rounded-lg text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition"
                >
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/>
                  </svg>
                </button>
              </>
            ) : (
              <button
                onClick={() => setShowLogin(true)}
                className="flex items-center gap-1.5 text-xs font-medium text-white bg-blue-600 hover:bg-blue-500 px-3 py-1.5 rounded-lg transition shadow-sm"
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/>
                </svg>
                登录
              </button>
            )}
          </div>
        </div>
      </header>

      <div className="flex flex-1 overflow-hidden relative">
        {/* Mobile sidebar backdrop */}
        {!sidebarCollapsed && (
          <div
            className="md:hidden fixed inset-0 bg-black/40 z-30"
            onClick={() => setSidebarCollapsed(true)}
          />
        )}

        {/* Sidebar */}
        <aside className={`
          ${sidebarCollapsed ? "w-0 max-md:-translate-x-full" : "w-64 max-md:translate-x-0"}
          border-r border-zinc-200 dark:border-zinc-800 shrink-0 bg-white dark:bg-zinc-900/95 backdrop-blur-sm flex flex-col
          transition-all duration-300 ease-in-out overflow-hidden relative
          max-md:fixed max-md:inset-y-0 max-md:left-0 max-md:z-40 max-md:shadow-xl
        `}>
          <div className="p-3 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between shrink-0">
            <span className="text-xs text-zinc-500 font-medium uppercase tracking-wider whitespace-nowrap">存储列表</span>
            <div className="flex items-center gap-2">
              {isAdmin && (
                <button
                  onClick={() => { setEditingStorage(null); setShowStorageForm(true); }}
                  className="inline-flex items-center gap-1 text-xs text-blue-500 hover:text-blue-400 font-medium px-2 py-1 rounded-md hover:bg-blue-50 dark:hover:bg-blue-900/20 transition whitespace-nowrap"
                >
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  添加
                </button>
              )}
              <button
                onClick={() => setSidebarCollapsed(true)}
                className="text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-300 p-1 rounded hover:bg-zinc-100 dark:hover:bg-zinc-800 transition"
                title="收起侧边栏"
              >
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><polyline points="15 18 9 12 15 6"/></svg>
              </button>
            </div>
          </div>
          <div className="overflow-y-auto flex-1">
            {storages.length === 0 ? (
              <div className="p-6 text-center">
                <svg className="mx-auto mb-2 text-zinc-300 dark:text-zinc-600" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
                  <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
                </svg>
                <p className="text-xs text-zinc-400 dark:text-zinc-600 font-medium">暂无存储</p>
                {isAdmin && <p className="text-xs text-zinc-400 dark:text-zinc-600 mt-1">点击上方「添加」创建</p>}
              </div>
            ) : (
              storages.map((s) => (
                <div
                  key={s.id}
                  className={`group flex items-center justify-between px-3 py-2.5 cursor-pointer border-l-2 transition-all ${
                    selectedStorage?.id === s.id
                      ? "border-blue-500 bg-blue-50/80 dark:bg-blue-900/20"
                      : "border-transparent hover:bg-zinc-50 dark:hover:bg-zinc-800/40"
                  }`}
                  onClick={() => setSelectedStorage(s)}
                >
                  <div className="flex items-center gap-2.5 min-w-0">
                    {/* Storage type icon */}
                    <div className={`shrink-0 w-7 h-7 rounded-md flex items-center justify-center ${
                      selectedStorage?.id === s.id
                        ? "bg-blue-100 dark:bg-blue-900/30"
                        : "bg-zinc-100 dark:bg-zinc-800"
                    }`}>
                      {s.type === "onedrive" ? (
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" className="text-blue-500"><path d="M20 18.5a3.5 3.5 0 0 0-1.5-2.89A5.006 5.006 0 0 0 6.5 11a4 4 0 0 0-.5 7.97A3.502 3.502 0 0 0 10 22h8.5a3.5 3.5 0 0 0 1.5-3.5z"/></svg>
                      ) : s.type === "gdrive" ? (
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="text-green-500"><path d="M4 20h16M4 20l4-16h8l4 16M7.5 8h9"/></svg>
                      ) : s.type === "alicloud" ? (
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" className="text-orange-500"><circle cx="12" cy="12" r="10"/><path d="M8 12h8M12 8v8"/></svg>
                      ) : s.type === "baiduyun" ? (
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" className="text-blue-400"><path d="M22 12h-4l-3 9L9 3l-3 9H2"/></svg>
                      ) : s.type === "webdev" ? (
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" className="text-purple-500"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>
                      ) : (
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" className="text-blue-500"><path d="M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z"/><line x1="4" y1="22" x2="4" y2="15"/></svg>
                      )}
                    </div>
                    <div className="min-w-0">
                      <div className={`text-sm font-medium truncate ${selectedStorage?.id === s.id ? "text-zinc-900 dark:text-zinc-100" : "text-zinc-600 dark:text-zinc-400"}`}>
                        {s.name}
                      </div>
                      <div className="flex items-center gap-1.5 mt-0.5">
                        <span className={`inline-block w-1.5 h-1.5 rounded-full ${s.isPublic ? "bg-green-400" : "bg-zinc-300 dark:bg-zinc-600"}`}></span>
                        <span className="text-[11px] text-zinc-400 dark:text-zinc-600 whitespace-nowrap">
                          {s.isPublic ? "公开" : "私有"}
                        </span>
                      </div>
                    </div>
                  </div>
                  {isAdmin && (
                    <div
                      className="flex items-center gap-0.5 opacity-100 md:opacity-0 md:group-hover:opacity-100 transition-opacity"
                      onClick={(e) => e.stopPropagation()}
                    >
                      <button
                        onClick={() => { setStatsStorage(s); setShowStats(true); }}
                        className="w-7 h-7 lg:w-auto lg:h-auto lg:px-2 lg:py-1 flex items-center justify-center rounded lg:rounded-md text-zinc-400 dark:text-zinc-500 hover:text-blue-500 dark:hover:text-blue-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 lg:border lg:border-zinc-200/70 lg:bg-white/80 lg:shadow-sm lg:hover:border-blue-300 lg:hover:bg-blue-50 lg:dark:border-zinc-700/70 lg:dark:bg-zinc-900/80 lg:dark:hover:border-blue-400/40 lg:dark:hover:bg-blue-500/10 lg:dark:hover:text-blue-300 lg:gap-1 transition"
                        title="统计"
                        aria-label="统计"
                      >
                        <svg className="lg:hidden" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>
                        <StatsIcon className="hidden lg:block h-4 w-4" />
                      </button>
                      <button
                        onClick={() => { setEditingStorage(s); setShowStorageForm(true); }}
                        className="w-7 h-7 flex items-center justify-center rounded text-zinc-400 dark:text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition"
                        title="编辑"
                      >
                        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                      <button
                        onClick={() => handleDeleteStorage(s)}
                        className="w-7 h-7 flex items-center justify-center rounded text-zinc-400 dark:text-zinc-500 hover:text-red-500 dark:hover:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20 transition"
                        title="删除"
                      >
                        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                      </button>
                    </div>
                  )}
                </div>
              ))
            )}
          </div>
        </aside>

        {/* Sidebar Expand Button - only show when collapsed on desktop */}
        {sidebarCollapsed && (
          <button
            onClick={() => setSidebarCollapsed(false)}
            className="hidden md:flex absolute left-0 top-1/2 -translate-y-1/2 z-10 w-5 h-10 items-center justify-center bg-white dark:bg-zinc-800 hover:bg-zinc-50 dark:hover:bg-zinc-700 text-zinc-400 dark:text-zinc-500 rounded-r-lg border border-l-0 border-zinc-200 dark:border-zinc-700 shadow-sm transition-colors"
            title="展开侧边栏"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><polyline points="9 18 15 12 9 6"/></svg>
          </button>
        )}

        {/* Main */}
        <main className="flex-1 bg-zinc-50 dark:bg-zinc-900 min-w-0 overflow-hidden">
          {selectedStorage ? (
            <FileBrowser storage={selectedStorage} isAdmin={isAdmin} isDark={isDark} chunkSizeMB={chunkSizeMB} />
          ) : (
            <div className="flex flex-col items-center justify-center h-full gap-3 px-4">
              <div className="w-16 h-16 rounded-2xl bg-zinc-100 dark:bg-zinc-800 flex items-center justify-center">
                <svg className="text-zinc-300 dark:text-zinc-600" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round">
                  <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
                  <line x1="12" y1="11" x2="12" y2="17"/><line x1="9" y1="14" x2="15" y2="14"/>
                </svg>
              </div>
              <div className="text-center">
                <p className="text-sm font-medium text-zinc-500 dark:text-zinc-400">选择一个存储</p>
                <p className="text-xs text-zinc-400 dark:text-zinc-600 mt-1">从左侧列表中选择存储以浏览文件</p>
              </div>
            </div>
          )}
        </main>
      </div>

      {/* Footer */}
      <footer className="shrink-0 border-t border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 px-4 py-2">
        <div className="flex items-center justify-center gap-3 md:gap-4 text-xs text-zinc-500 dark:text-zinc-500">
          <a
            href="https://github.com/ooyyh/Cloudflare-Clist"
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-1 hover:text-zinc-700 dark:hover:text-zinc-300 transition"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M12 0C5.37 0 0 5.37 0 12c0 5.31 3.435 9.795 8.205 11.385.6.105.825-.255.825-.57 0-.285-.015-1.23-.015-2.235-3.015.555-3.795-.735-4.035-1.41-.135-.345-.72-1.41-1.23-1.695-.42-.225-1.02-.78-.015-.795.945-.015 1.62.87 1.845 1.23 1.08 1.815 2.805 1.305 3.495.99.105-.78.42-1.305.765-1.605-2.67-.3-5.46-1.335-5.46-5.925 0-1.305.465-2.385 1.23-3.225-.12-.3-.54-1.53.12-3.18 0 0 1.005-.315 3.3 1.23.96-.27 1.98-.405 3-.405s2.04.135 3 .405c2.295-1.56 3.3-1.23 3.3-1.23.66 1.65.24 2.88.12 3.18.765.84 1.23 1.905 1.23 3.225 0 4.605-2.805 5.625-5.475 5.925.435.375.81 1.095.81 2.22 0 1.605-.015 2.895-.015 3.3 0 .315.225.69.825.57A12.02 12.02 0 0 0 24 12c0-6.63-5.37-12-12-12z"/></svg>
            <span className="hidden sm:inline">GitHub</span>
          </a>
          <span className="text-zinc-200 dark:text-zinc-700">·</span>
          <button
            onClick={() => setShowChangelog(true)}
            className="inline-flex items-center gap-1 hover:text-zinc-700 dark:hover:text-zinc-300 transition"
          >
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
            更新日志
          </button>
          <span className="text-zinc-200 dark:text-zinc-700 hidden sm:inline">·</span>
          <span className="hidden sm:inline">Made by <span className="text-zinc-600 dark:text-zinc-400 font-medium">ooyyh</span></span>
          <span className="text-zinc-200 dark:text-zinc-700 hidden sm:inline">·</span>
          <span className="hidden sm:inline-flex items-center gap-1">
            Powered by
            <a
              href="https://www.cloudflare.com"
              target="_blank"
              rel="noopener noreferrer"
              className="text-orange-500 hover:text-orange-400 transition font-medium"
            >
              Cloudflare
            </a>
          </span>
        </div>
      </footer>

      {/* Modals */}
      {showLogin && (
        <LoginModal
          onLogin={() => { setShowLogin(false); refreshStorages(); setIsAdmin(true); }}
          onClose={() => setShowLogin(false)}
        />
      )}
      {showStorageForm && (
        <StorageModal
          storage={editingStorage || undefined}
          onSave={() => { setShowStorageForm(false); setEditingStorage(null); refreshStorages(); }}
          onCancel={() => { setShowStorageForm(false); setEditingStorage(null); }}
        />
      )}
      {showSettings && (
        <SettingsModal
          onClose={() => setShowSettings(false)}
          siteTitle={siteTitle}
          siteAnnouncement={siteAnnouncement}
          isDark={isDark}
          onToggleTheme={toggleTheme}
          isAdmin={isAdmin}
          onRefreshStorages={refreshStorages}
          webdavEnabled={webdavEnabled}
          storages={storages}
        />
      )}
      {showAnnouncement && siteAnnouncement && (
        <AnnouncementModal
          announcement={siteAnnouncement}
          onClose={() => setShowAnnouncement(false)}
        />
      )}
      {showChangelog && (
        <ChangelogModal onClose={() => setShowChangelog(false)} />
      )}
      {showStats && statsStorage && (
        <StorageStatsModal
          storage={statsStorage}
          onClose={() => { setShowStats(false); setStatsStorage(null); }}
        />
      )}
    </div>
  );
}
