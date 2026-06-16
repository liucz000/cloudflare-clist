import type { Route } from "./+types/share";

export async function loader({ request }: Route.LoaderArgs) {
  const url = new URL(request.url);
  const token = url.searchParams.get("token");

  if (!token) {
    return { token: null };
  }

  return { token };
}

import { useState, useEffect } from "react";

interface S3Object {
  key: string;
  name: string;
  size: number;
  lastModified: string;
  isDirectory: boolean;
}

interface Share {
  id: string;
  storageId: number;
  filePath: string;
  isDirectory: boolean;
  shareToken: string;
  expiresAt: string | null;
  createdAt: string;
}

interface StorageInfo {
  id: number;
  name: string;
}

function formatBytes(bytes: number): string {
  if (bytes === 0) return "-";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB", "TB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

function formatDate(dateStr: string): string {
  if (!dateStr) return "-";
  const date = new Date(dateStr);
  return date.toLocaleString("zh-CN");
}

function getFileIcon(name: string): string {
  const ext = name.split(".").pop()?.toLowerCase() || "";
  const iconMap: Record<string, string> = {
    pdf: "📄",
    doc: "📝",
    docx: "📝",
    xls: "📊",
    xlsx: "📊",
    ppt: "🎬",
    pptx: "🎬",
    zip: "📦",
    rar: "📦",
    "7z": "📦",
    jpg: "🖼",
    jpeg: "🖼",
    png: "🖼",
    gif: "🖼",
    mp4: "🎥",
    avi: "🎥",
    mkv: "🎥",
    mp3: "🎵",
    flac: "🎵",
    wav: "🎵",
    txt: "📃",
    json: "⚙️",
    xml: "⚙️",
    yaml: "⚙️",
    yml: "⚙️",
  };
  return iconMap[ext] || "📄";
}

export default function Share({ loaderData }: Route.ComponentProps) {
  const { token } = loaderData as { token: string | null };
  const [share, setShare] = useState<Share | null>(null);
  const [storage, setStorage] = useState<StorageInfo | null>(null);
  const [path, setPath] = useState("");
  const [objects, setObjects] = useState<S3Object[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    if (!token) {
      setError("分享令牌缺失");
      setLoading(false);
      return;
    }

    const fetchShareInfo = async () => {
      try {
        const res = await fetch(`/api/shares?token=${token}`);
        if (res.ok) {
          const data = (await res.json()) as { share: Share; storage: StorageInfo };
          setShare(data.share);
          setStorage(data.storage);
        } else {
          const data = (await res.json()) as { error?: string };
          setError(data.error || "分享链接不存在或已过期");
        }
      } catch {
        setError("网络错误");
      } finally {
        setLoading(false);
      }
    };

    fetchShareInfo();
  }, [token]);

  useEffect(() => {
    if (share && storage) {
      loadFiles();
    }
  }, [share, storage, path]);

  const loadFiles = async () => {
    if (!share || !storage || !token) return;

    setLoading(true);

    try {
      // If sharing a single file and at root level, show the file itself
      if (!share.isDirectory && !path) {
        const fileName = share.filePath.split("/").pop() || share.filePath;
        const fileObj: S3Object = {
          key: share.filePath,
          name: fileName,
          size: 0,
          lastModified: share.createdAt,
          isDirectory: false,
        };
        setObjects([fileObj]);
        setLoading(false);
        return;
      }

      const basePath = share.filePath;
      const fullPath = path ? `${basePath}/${path}` : basePath;

      const res = await fetch(
        `/api/files/${storage.id}/${fullPath}?action=list&token=${token}`
      );

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
    // If sharing a single file, don't allow navigation
    if (!share?.isDirectory) return;
    setPath(newPath.replace(/^\//, "").replace(/\/$/, ""));
  };

  const downloadFile = (key: string) => {
    window.open(
      `/api/files/${storage!.id}/${key}?action=download&token=${token}`,
      "_blank"
    );
  };

  if (error && !share) {
    return (
      <div className="flex items-center justify-center h-screen bg-white dark:bg-zinc-950">
        <div className="text-center p-8 bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg">
          <div className="text-red-500 dark:text-red-400 font-mono text-lg">
            ❌ {error}
          </div>
        </div>
      </div>
    );
  }

  if (!share || !storage) {
    return (
      <div className="flex items-center justify-center h-screen bg-white dark:bg-zinc-950">
        <div className="text-center p-8">
          <div className="text-zinc-500 dark:text-zinc-400 font-mono">
            加载中...
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="h-screen w-full flex flex-col bg-white dark:bg-zinc-950 text-zinc-900 dark:text-zinc-100">
      {/* Header */}
      <div className="px-5 py-3 border-b border-zinc-200 dark:border-zinc-800 shrink-0">
        <div className="flex items-center justify-between">
          <div>
            <div className="text-base font-mono font-semibold flex items-center gap-2">🔗 分享内容</div>
            <div className="text-xs text-zinc-500 font-mono mt-0.5">
              {storage.name} / {share.filePath}
              {share.expiresAt && (
                <span className="ml-2 text-yellow-600 dark:text-yellow-400">
                  ⏰ {formatDate(share.expiresAt)}
                </span>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Breadcrumb */}
      <div className="px-5 py-2 border-b border-zinc-200 dark:border-zinc-800 shrink-0 flex items-center gap-1.5 text-xs font-mono">
        <button
          onClick={() => setPath("")}
          className="text-blue-500 hover:text-blue-400 transition-colors"
        >
          📁 根目录
        </button>
        {path
          .split("/")
          .filter(Boolean)
          .map((part, index, arr) => {
            const fullPath = arr.slice(0, index + 1).join("/");
            return (
              <div key={fullPath} className="flex items-center gap-1.5">
                <span className="text-zinc-300 dark:text-zinc-600">/</span>
                <button
                  onClick={() => navigateTo(fullPath)}
                  className="text-blue-500 hover:text-blue-400 transition-colors"
                >
                  {part}
                </button>
              </div>
            );
          })}
      </div>

      {/* Content */}
      <div className="flex-1 overflow-auto">
        {loading ? (
          <div className="flex items-center justify-center h-32 text-zinc-500 font-mono text-sm">
            加载中...
          </div>
        ) : error ? (
          <div className="flex items-center justify-center h-32 text-red-500 dark:text-red-400 font-mono text-sm">
            {error}
          </div>
        ) : objects.length === 0 ? (
          <div className="flex items-center justify-center h-32 text-zinc-400 dark:text-zinc-600 font-mono text-sm">
            空目录
          </div>
        ) : (
          <div className="divide-y divide-zinc-100 dark:divide-zinc-800/60">
            {/* List header */}
            <div className="flex items-center gap-4 px-5 py-1.5 text-xs text-zinc-400 dark:text-zinc-500 font-mono bg-zinc-50 dark:bg-zinc-900/80">
              <div className="flex-1">名称</div>
              <div className="flex items-center gap-4 shrink-0">
                <span className="hidden sm:inline w-40 text-left">修改时间</span>
                <span className="w-14 text-center">操作</span>
              </div>
            </div>
            {objects.map((obj) => (
              <div
                key={obj.key}
                className="flex items-center gap-3 px-5 py-2.5 hover:bg-zinc-50 dark:hover:bg-zinc-900/50 transition-colors group"
              >
                {/* Icon + Name */}
                <div className="flex items-center gap-2.5 min-w-0 flex-1">
                  {obj.isDirectory ? (
                    <button
                      onClick={() => navigateTo(obj.key)}
                      className="flex items-center gap-2.5 min-w-0 text-blue-500 hover:text-blue-400 transition-colors"
                    >
                      <span className="text-lg leading-none shrink-0">📁</span>
                      <span className="font-mono text-sm truncate">{obj.name}</span>
                    </button>
                  ) : (
                    <>
                      <span className="text-lg leading-none shrink-0">{getFileIcon(obj.name)}</span>
                      <span className="font-mono text-sm text-zinc-700 dark:text-zinc-300 truncate">{obj.name}</span>
                    </>
                  )}
                </div>

                {/* Meta: date · action — compact, left-aligned */}
                <div className="flex items-center gap-4 shrink-0 text-xs font-mono">
                  {!obj.isDirectory && (
                    <>
                      <span className="text-zinc-500 dark:text-zinc-400 hidden sm:inline w-40 text-left">
                        {formatDate(obj.lastModified)}
                      </span>
                      <button
                        onClick={() => downloadFile(obj.key)}
                        className="flex items-center gap-1 px-2.5 py-1 bg-zinc-100 dark:bg-zinc-800 hover:bg-blue-500 dark:hover:bg-blue-500 text-zinc-500 dark:text-zinc-400 hover:text-white rounded transition-colors"
                        title="下载"
                      >
                        <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
                          <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                        </svg>
                        <span className="hidden sm:inline">下载</span>
                      </button>
                    </>
                  )}
                  {obj.isDirectory && (
                    <span className="text-zinc-400 dark:text-zinc-500">-</span>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Footer */}
      <div className="px-5 py-3 border-t border-zinc-200 dark:border-zinc-800 shrink-0 text-xs text-zinc-400 dark:text-zinc-600 font-mono">
        CList 分享
      </div>
    </div>
  );
}
