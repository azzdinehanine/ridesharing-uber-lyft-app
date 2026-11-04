package com.mindorks.ridesharing.ui.login

import android.content.Intent
import android.os.Bundle
import android.util.Log
import android.view.View
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import com.microsoft.identity.client.AuthenticationCallback
import com.microsoft.identity.client.IAccount
import com.microsoft.identity.client.IAuthenticationResult
import com.microsoft.identity.client.ISingleAccountPublicClientApplication
import com.microsoft.identity.client.exception.MsalException
import com.mindorks.ridesharing.R
import com.mindorks.ridesharing.auth.OutlookAuthManager
import com.mindorks.ridesharing.ui.maps.MapsActivity
import kotlinx.android.synthetic.main.activity_login.*

class LoginActivity : AppCompatActivity() {

    companion object {
        private const val TAG = "LoginActivity"
    }

    private var msalApp: ISingleAccountPublicClientApplication? = null

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_login)

        outlookSignInButton.setOnClickListener { signIn() }

        if (!OutlookAuthManager.isConfigured()) {
            showMessage(getString(R.string.sign_in_not_configured))
            outlookSignInButton.isEnabled = false
            return
        }

        setLoading(true)
        OutlookAuthManager.getApplication(this, { app ->
            msalApp = app
            checkExistingAccount(app)
        }, { error ->
            Log.e(TAG, "Unable to create MSAL application", error)
            setLoading(false)
            showMessage(getString(R.string.sign_in_failed, error.message))
        })
    }

    /** Skip the login screen if the user already signed in with Outlook earlier. */
    private fun checkExistingAccount(app: ISingleAccountPublicClientApplication) {
        app.getCurrentAccountAsync(object :
            ISingleAccountPublicClientApplication.CurrentAccountCallback {
            override fun onAccountLoaded(activeAccount: IAccount?) {
                setLoading(false)
                if (activeAccount != null) openMaps(activeAccount.username)
            }

            override fun onAccountChanged(priorAccount: IAccount?, currentAccount: IAccount?) {
                setLoading(false)
                if (currentAccount != null) openMaps(currentAccount.username)
            }

            override fun onError(exception: MsalException) {
                Log.e(TAG, "Unable to load current account", exception)
                setLoading(false)
            }
        })
    }

    private fun signIn() {
        val app = msalApp ?: return
        setLoading(true)
        app.signIn(this, null, OutlookAuthManager.SCOPES, object : AuthenticationCallback {
            override fun onSuccess(authenticationResult: IAuthenticationResult) {
                setLoading(false)
                openMaps(authenticationResult.account.username)
            }

            override fun onError(exception: MsalException) {
                Log.e(TAG, "Outlook sign in failed", exception)
                setLoading(false)
                showMessage(getString(R.string.sign_in_failed, exception.message))
            }

            override fun onCancel() {
                setLoading(false)
                showMessage(getString(R.string.sign_in_cancelled))
            }
        })
    }

    private fun openMaps(email: String?) {
        if (!email.isNullOrBlank()) {
            Toast.makeText(this, getString(R.string.signed_in_as, email), Toast.LENGTH_SHORT).show()
        }
        startActivity(Intent(this, MapsActivity::class.java))
        finish()
    }

    private fun setLoading(loading: Boolean) {
        runOnUiThread {
            loginProgressBar.visibility = if (loading) View.VISIBLE else View.GONE
            outlookSignInButton.isEnabled = !loading && msalApp != null
        }
    }

    private fun showMessage(message: String) {
        runOnUiThread {
            loginMessageTextView.visibility = View.VISIBLE
            loginMessageTextView.text = message
        }
    }
}
