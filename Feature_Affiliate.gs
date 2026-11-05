/**
 * Affiliate share link (TimPRO) — weekly review.
 * Sheet "Affiliate (Timpro)": daily funnel app DAU → open affiliate dashboard → share affiliate link.
 */

function affiliateFeature_() {
  return createFunnelFeature_({
    id: 'affiliate',
    name: 'Affiliate',
    menuFunction: 'generateAffiliateSlides',
    sheet: 'Affiliate (Timpro)',
    description: 'Pengurus share a referral link so other people log in and buy at the Toko Sawit store; the pengurus earns a ' +
      'commission on those purchases. Funnel: open the affiliate dashboard → share the affiliate link.',
    app: 'timpro',
    fields: {
      date: ['date', 'tanggal'],
      dau: ['step1_dau_app', 'app_dau', 'daily_active_users'],
      dashboard: ['step2_open_affiliate_dashboard'],
      share: ['step3_share_affiliate_link'],
      newUsers: ['step2_new_users'],
      returningUsers: ['step2_returning_users'],
      names: ['step2_usernames', 'step2_user_names'],
      appWau: ['weekly_app_wau', 'app_weekly_active_user'],
      featureWau: ['weekly_feature_wau', 'feature_weekly_active_user']
    },
    steps: [
      { key: 'dashboard', label: 'Open affiliate dashboard', short: 'Dashboard' },
      { key: 'share', label: 'Share affiliate link', short: 'Share link' }
    ],
    usersLabel: 'USERS OPEN AFFILIATE DASHBOARD'
  });
}

function generateAffiliateSlides() {
  generateFeatures_(['affiliate']);
}
