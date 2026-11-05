/**
 * Lapak Sawit — weekly review, one slide per app (AgenPRO, TimPRO).
 * Sheet "Lapak Sawit": one row per date per app (feature_name "Lapak Sawit Agenpro" / "Lapak Sawit Timpro").
 */

function lapakSawitFeature_() {
  return createFunnelFeature_({
    id: 'lapak_sawit',
    name: 'Lapak Sawit',
    menuFunction: 'generateLapakSawitSlides',
    sheet: 'Lapak Sawit',
    description: 'Marketplace connecting agents and farmers (petani) to run FFB/TBS transactions. Available in both AgenPRO and TimPRO.',
    splitByColumn: true,
    fields: {
      date: ['tanggal', 'date'],
      segment: ['feature_name', 'feature'],
      dau: ['daily_active_users', 'dau', 'app_dau'],
      feature: ['feature_lapaksawit_users', 'feature_unique_users', 'feature_users', /^feature_(?!weekly).+_users$/],
      newUsers: ['new_users'],
      returningUsers: ['returning_user', 'returning_users'],
      names: ['user_names', 'usernames'],
      appWau: ['app_weekly_active_user', 'app_weekly_active_users', 'weekly_app_wau', 'app_wau'],
      featureWau: ['feature_weekly_active_user', 'feature_weekly_active_users', 'weekly_feature_wau', 'feature_wau']
    },
    steps: [{ key: 'feature', label: 'Klik Lapak Sawit' }],
    usersLabel: 'USERS KLIK LAPAK SAWIT'
  });
}

function generateLapakSawitSlides() {
  generateFeatures_(['lapak_sawit']);
}
