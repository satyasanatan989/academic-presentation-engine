package com.academic.presentationengine;

import android.annotation.SuppressLint;
import android.content.ContentValues;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.content.pm.ResolveInfo;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.util.Base64;
import android.util.Log;
import android.view.View;
import android.webkit.ConsoleMessage;
import android.webkit.JavascriptInterface;
import android.webkit.JsResult;
import android.webkit.WebChromeClient;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Toast;
import androidx.appcompat.app.AlertDialog;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.FileProvider;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public class MainActivity extends AppCompatActivity {

    private static final String TAG = "ScholarDeck";
    private WebView webView;

    public class WebAppInterface {
        Context mContext;

        WebAppInterface(Context c) {
            mContext = c;
        }

        @JavascriptInterface
        public void saveAndOpenPptxBase64(String fileName, String base64Data) {
            runOnUiThread(() -> {
                try {
                    Log.d(TAG, "Saving PPTX Base64: " + fileName);
                    byte[] bytes = Base64.decode(base64Data, Base64.DEFAULT);
                    
                    // 1. Save to App-Specific External Files Dir (100% Guaranteed Permission-Free)
                    File appDownloads = mContext.getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS);
                    if (appDownloads != null && !appDownloads.exists()) {
                        appDownloads.mkdirs();
                    }
                    File pptFile = new File(appDownloads, fileName);
                    FileOutputStream fos = new FileOutputStream(pptFile);
                    fos.write(bytes);
                    fos.flush();
                    fos.close();
                    Log.d(TAG, "Saved to App Storage: " + pptFile.getAbsolutePath() + " (" + bytes.length + " bytes)");

                    // 2. Also export to Public Downloads via MediaStore on Android 10+
                    if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                        try {
                            ContentValues values = new ContentValues();
                            values.put(android.provider.MediaStore.MediaColumns.DISPLAY_NAME, fileName);
                            values.put(android.provider.MediaStore.MediaColumns.MIME_TYPE, "application/vnd.openxmlformats-officedocument.presentationml.presentation");
                            values.put(android.provider.MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS);
                            Uri mediaUri = mContext.getContentResolver().insert(android.provider.MediaStore.Downloads.EXTERNAL_CONTENT_URI, values);
                            if (mediaUri != null) {
                                OutputStream os = mContext.getContentResolver().openOutputStream(mediaUri);
                                if (os != null) {
                                    os.write(bytes);
                                    os.flush();
                                    os.close();
                                    Log.d(TAG, "Exported to Public Downloads via MediaStore: " + mediaUri);
                                }
                            }
                        } catch (Exception ex) {
                            Log.w(TAG, "MediaStore export warning: " + ex.getMessage());
                        }
                    }

                    Toast.makeText(mContext, "Presentation Ready: " + fileName, Toast.LENGTH_SHORT).show();

                    // 3. Immediately launch in Microsoft PowerPoint / Google Slides / WPS Office
                    launchPptFile(pptFile);

                } catch (Exception e) {
                    Log.e(TAG, "Error saving PPT: " + e.getMessage(), e);
                    Toast.makeText(mContext, "Error saving PPT: " + e.getMessage(), Toast.LENGTH_LONG).show();
                }
            });
        }

        @JavascriptInterface
        public void openPptx(String fileName) {
            openPptxFile(fileName);
        }

        @JavascriptInterface
        public void openPptxFile(String fileName) {
            runOnUiThread(() -> {
                try {
                    Log.d(TAG, "Opening PPTX: " + fileName);
                    File appDownloads = mContext.getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS);
                    File pptFile = new File(appDownloads, fileName);

                    // If file not in app storage, check public downloads
                    if (!pptFile.exists()) {
                        File pubFile = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), fileName);
                        if (pubFile.exists()) {
                            pptFile = pubFile;
                        }
                    }

                    // If still not found, check APK assets/decks/
                    if (!pptFile.exists() || pptFile.length() == 0) {
                        try {
                            InputStream is = mContext.getAssets().open("decks/" + fileName);
                            pptFile = new File(appDownloads, fileName);
                            FileOutputStream fos = new FileOutputStream(pptFile);
                            byte[] buffer = new byte[8192];
                            int read;
                            while ((read = is.read(buffer)) != -1) {
                                fos.write(buffer, 0, read);
                            }
                            fos.flush();
                            fos.close();
                            is.close();
                            Log.d(TAG, "Extracted deck from assets: " + fileName + " (" + pptFile.length() + " bytes)");
                        } catch (Exception assetEx) {
                            Log.d(TAG, "Not in assets: " + assetEx.getMessage());
                        }
                    }

                    if (pptFile.exists() && pptFile.length() > 0) {
                        launchPptFile(pptFile);
                    } else {
                        Toast.makeText(mContext, "Presentation not found: " + fileName, Toast.LENGTH_LONG).show();
                    }

                } catch (Exception e) {
                    Log.e(TAG, "Could not open PPT: " + e.getMessage(), e);
                    Toast.makeText(mContext, "Could not open PPT: " + e.getMessage(), Toast.LENGTH_SHORT).show();
                }
            });
        }

        private void launchPptFile(File file) {
            try {
                Uri uri = FileProvider.getUriForFile(mContext, mContext.getPackageName() + ".provider", file);
                Intent intent = new Intent(Intent.ACTION_VIEW);
                intent.setDataAndType(uri, "application/vnd.openxmlformats-officedocument.presentationml.presentation");
                intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);

                // Grant URI permission explicitly to any resolving app
                List<ResolveInfo> resInfoList = mContext.getPackageManager().queryIntentActivities(intent, PackageManager.MATCH_DEFAULT_ONLY);
                for (ResolveInfo resolveInfo : resInfoList) {
                    mContext.grantUriPermission(resolveInfo.activityInfo.packageName, uri, Intent.FLAG_GRANT_READ_URI_PERMISSION);
                }

                Intent chooser = Intent.createChooser(intent, "Open Presentation With");
                chooser.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
                mContext.startActivity(chooser);
            } catch (Exception e) {
                Log.e(TAG, "Launch chooser failed: " + e.getMessage(), e);
                Toast.makeText(mContext, "No PowerPoint viewer found. Please install PowerPoint or Google Slides.", Toast.LENGTH_LONG).show();
            }
        }

        @JavascriptInterface
        public void showToast(String message) {
            runOnUiThread(() -> Toast.makeText(mContext, message, Toast.LENGTH_SHORT).show());
        }
    }

    @SuppressLint("SetJavaScriptEnabled")
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        getWindow().getDecorView().setSystemUiVisibility(
            View.SYSTEM_UI_FLAG_LAYOUT_STABLE
        );

        webView = new WebView(this);
        setContentView(webView);

        WebSettings settings = webView.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        settings.setDatabaseEnabled(true);
        settings.setAllowFileAccess(true);
        settings.setAllowContentAccess(true);
        settings.setLoadWithOverviewMode(true);
        settings.setUseWideViewPort(true);
        settings.setBuiltInZoomControls(false);
        settings.setDisplayZoomControls(false);
        settings.setCacheMode(WebSettings.LOAD_DEFAULT);
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.LOLLIPOP) {
            settings.setMixedContentMode(WebSettings.MIXED_CONTENT_ALWAYS_ALLOW);
        }

        webView.addJavascriptInterface(new WebAppInterface(this), "AndroidHost");

        webView.setWebViewClient(new WebViewClient());
        webView.setWebChromeClient(new WebChromeClient() {
            @Override
            public boolean onConsoleMessage(ConsoleMessage cm) {
                Log.d(TAG + "-JS", cm.message() + " -- From line " + cm.lineNumber() + " of " + cm.sourceId());
                return true;
            }

            @Override
            public boolean onJsAlert(WebView view, String url, String message, JsResult result) {
                new AlertDialog.Builder(MainActivity.this)
                    .setTitle("ScholarDeck")
                    .setMessage(message)
                    .setPositiveButton(android.R.string.ok, (dialog, which) -> result.confirm())
                    .setCancelable(false)
                    .create()
                    .show();
                return true;
            }
        });

        webView.loadUrl("file:///android_asset/index.html");
    }

    @Override
    public void onBackPressed() {
        if (webView != null && webView.canGoBack()) {
            webView.goBack();
        } else {
            super.onBackPressed();
        }
    }
}
