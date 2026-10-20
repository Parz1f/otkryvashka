package app.otkryvashka.rust

import android.content.Intent
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.provider.OpenableColumns
import android.webkit.WebView
import org.json.JSONObject
import java.io.File
import kotlin.concurrent.thread

class MainActivity : TauriActivity() {
  private var webView: WebView? = null

  override fun onCreate(savedInstanceState: Bundle?) {
    super.onCreate(savedInstanceState)
    if (savedInstanceState == null) receive(intent)
  }

  override fun onWebViewCreate(webView: WebView) {
    this.webView = webView
  }

  override fun onNewIntent(intent: Intent) {
    super.onNewIntent(intent)
    receive(intent)
  }

  /**
   * Файл из «Открыть с помощью» или «Поделиться» копируется в кэш (cache/incoming),
   * откуда его читает Rust (команды incoming_info и take_incoming). Описание файла
   * пишется последним, поэтому Rust никогда не увидит недокопированный файл.
   */
  private fun receive(intent: Intent?) {
    if (intent == null) return
    val uri: Uri = when (intent.action) {
      Intent.ACTION_VIEW -> intent.data
      Intent.ACTION_SEND ->
        if (Build.VERSION.SDK_INT >= 33) intent.getParcelableExtra(Intent.EXTRA_STREAM, Uri::class.java)
        else @Suppress("DEPRECATION") intent.getParcelableExtra(Intent.EXTRA_STREAM)
      else -> null
    } ?: return
    val mime = intent.type?.takeUnless { it.contains('*') } ?: contentResolver.getType(uri) ?: ""
    val name = displayName(uri)

    thread {
      val dir = File(cacheDir, "incoming").apply { mkdirs() }
      val meta = File(dir, "meta.json")
      val file = File(dir, "file")
      val part = File(dir, "file.part")
      meta.delete()
      try {
        val input = contentResolver.openInputStream(uri) ?: return@thread
        input.use { src -> part.outputStream().use { src.copyTo(it) } }
        file.delete()
        part.renameTo(file)
        val json = JSONObject().put("name", name).put("mime", mime).put("size", file.length())
        val metaPart = File(dir, "meta.part")
        metaPart.writeText(json.toString())
        metaPart.renameTo(meta)
        runOnUiThread { webView?.evaluateJavascript("window.checkIncoming && window.checkIncoming()", null) }
      } catch (e: Exception) {
        part.delete()
      }
    }
  }

  private fun displayName(uri: Uri): String {
    if (uri.scheme == "content") {
      try {
        contentResolver.query(uri, arrayOf(OpenableColumns.DISPLAY_NAME), null, null, null)?.use { c ->
          if (c.moveToFirst()) c.getString(0)?.takeIf { it.isNotEmpty() }?.let { return it }
        }
      } catch (_: Exception) {
      }
    }
    return uri.lastPathSegment?.substringAfterLast('/') ?: "Файл"
  }
}
