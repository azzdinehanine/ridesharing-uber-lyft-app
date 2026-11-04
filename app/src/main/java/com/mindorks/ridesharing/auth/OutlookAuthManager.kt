package com.mindorks.ridesharing.auth

import android.content.Context
import com.microsoft.identity.client.IPublicClientApplication
import com.microsoft.identity.client.ISingleAccountPublicClientApplication
import com.microsoft.identity.client.PublicClientApplication
import com.microsoft.identity.client.exception.MsalException
import com.mindorks.ridesharing.BuildConfig
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.net.URLEncoder

/**
 * Wraps MSAL (Microsoft Authentication Library) so the app can sign users in
 * with their Outlook.com / Hotmail / Live (personal) or Microsoft 365 (work/school) account.
 *
 * The client id and signature hash are read from local.properties at build time
 * (see app/build.gradle), so no secrets or per-developer values are committed.
 */
object OutlookAuthManager {

    /** Basic profile scope - enough to read the signed-in user's email address. */
    val SCOPES = arrayOf("User.Read")

    private var application: ISingleAccountPublicClientApplication? = null

    fun isConfigured(): Boolean =
        BuildConfig.MSAL_CLIENT_ID.isNotBlank() && BuildConfig.MSAL_SIGNATURE_HASH.isNotBlank()

    fun getApplication(
        context: Context,
        onReady: (ISingleAccountPublicClientApplication) -> Unit,
        onError: (Exception) -> Unit
    ) {
        application?.let {
            onReady(it)
            return
        }
        if (!isConfigured()) {
            onError(IllegalStateException("MSAL is not configured"))
            return
        }
        val configFile = writeConfigFile(context.applicationContext)
        PublicClientApplication.createSingleAccountPublicClientApplication(
            context.applicationContext,
            configFile,
            object : IPublicClientApplication.ISingleAccountApplicationCreatedListener {
                override fun onCreated(app: ISingleAccountPublicClientApplication) {
                    application = app
                    onReady(app)
                }

                override fun onError(exception: MsalException) {
                    onError(exception as Exception)
                }
            })
    }

    private fun redirectUri(): String {
        val encodedHash = URLEncoder.encode(BuildConfig.MSAL_SIGNATURE_HASH, "UTF-8")
        return "msauth://${BuildConfig.APPLICATION_ID}/$encodedHash"
    }

    private fun writeConfigFile(context: Context): File {
        val audience = JSONObject()
            // "common" + this audience type accepts personal Outlook.com/Hotmail/Live
            // accounts as well as work/school Microsoft 365 accounts.
            .put("type", "AzureADandPersonalMicrosoftAccount")
            .put("tenant_id", "common")
        val authority = JSONObject()
            .put("type", "AAD")
            .put("audience", audience)
        val config = JSONObject()
            .put("client_id", BuildConfig.MSAL_CLIENT_ID)
            .put("authorization_user_agent", "DEFAULT")
            .put("redirect_uri", redirectUri())
            .put("account_mode", "SINGLE")
            .put("broker_redirect_uri_registered", true)
            .put("authorities", JSONArray().put(authority))

        val file = File(context.filesDir, "msal_auth_config.json")
        file.writeText(config.toString())
        return file
    }
}
